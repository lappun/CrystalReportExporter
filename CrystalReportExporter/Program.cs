// =================================================================================
//  Crystal Reports Definition Extractor - FINAL COMPLETE SOURCE CODE
//  This version incorporates all fixes for typos, API errors, and ambiguities.
// =================================================================================
using System;
using System.IO;
using System.Linq;
using System.Text;

// Required using statements for all Crystal Reports types used in this file.
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using CrystalDecisions.ReportAppServer.ClientDoc;
using CrystalDecisions.ReportAppServer.Controllers;
using CrystalDecisions.ReportAppServer.DataDefModel;
using CrystalDecisions.ReportAppServer.CommonObjectModel;

namespace CrystalReportDocumenter
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {
            // --- 1. Argument Handling ---
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: CrystalReportDocumenter.exe \"<path_to_input.rpt>\" \"<path_to_output.txt>\"");
                return;
            }
            string reportPath = args[0];
            string outputPath = args[1];
            if (!File.Exists(reportPath))
            {
                Console.WriteLine($"Error: Input file not found at '{reportPath}'");
                return;
            }

            Console.WriteLine($"Processing: {Path.GetFileName(reportPath)}");
            var sb = new StringBuilder();
            ReportDocument reportDocument = new ReportDocument();

            try
            {
                // --- 2. Load the Report ---
                // FIX (Workaround): Use the simpler Load method that only takes the path.
                // This avoids the 'OpenReportMethod' enum which was causing errors in your environment.
                reportDocument.Load(reportPath);

                sb.AppendLine($"# REPORT DEFINITION: {Path.GetFileName(reportPath)}");
                sb.AppendLine($"# GENERATED ON: {DateTime.Now}");
                sb.AppendLine("====================================================================");

                // --- 3. Call Methods to Extract All Information ---
                ExtractDataSourceInfo(reportDocument, sb);
                ExtractParameters(reportDocument, sb);
                ExtractSelectionFormulas(reportDocument, sb);
                ExtractCustomFormulas(reportDocument, sb);
                ExtractGrouping(reportDocument, sb);

                // --- 4. Write to File ---
                File.WriteAllText(outputPath, sb.ToString());
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"\nSuccessfully created definition file: {outputPath}");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\nAn error occurred: {ex.Message}");
                Console.ResetColor();
                Console.WriteLine($"\nStackTrace:\n{ex.StackTrace}");
            }
            finally
            {
                // --- 5. Clean Up ---
                // This is critical to release file locks and memory.
                reportDocument.Close();
                reportDocument.Dispose();
            }
        }

        /// <summary>
        /// Extracts the data source connection info and lists the tables/procedures used.
        /// This version contains the fix for the invalid 'TableName' property.
        /// </summary>
        private static void ExtractDataSourceInfo(ReportDocument reportDocument, StringBuilder sb)
        {
            sb.AppendLine("\n--- DATA SOURCE ---");
            try
            {
                if (reportDocument.Database.Tables.Count == 0)
                {
                    sb.AppendLine("No database tables or procedures found in the report.");
                    return;
                }

                // FIX (Ambiguity): Use the full class name to be explicit.
                CrystalDecisions.CrystalReports.Engine.Table mainTable = reportDocument.Database.Tables[0];
                var connectionInfo = mainTable.LogOnInfo.ConnectionInfo;

                sb.AppendLine($"[Connection Type]: {connectionInfo.Type}");
                sb.AppendLine($"[Server Name]: {connectionInfo.ServerName}");
                sb.AppendLine($"[Database Name]: {connectionInfo.DatabaseName}");

                // FIX: A simpler, more reliable way to list the data sources.
                sb.AppendLine("\n[Sources Used (Tables, Views, or Stored Procedures)]:");
                foreach (CrystalDecisions.CrystalReports.Engine.Table table in reportDocument.Database.Tables)
                {
                    sb.AppendLine($"- {table.Name}");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Could not retrieve data source info. Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts parameter information.
        /// This version fixes the ambiguity error for 'ParameterField'.
        /// </summary>
        private static void ExtractParameters(ReportDocument reportDocument, StringBuilder sb)
        {
            sb.AppendLine("\n--- PARAMETERS ---");
            if (reportDocument.ParameterFields.Count == 0) { sb.AppendLine("No parameters found."); }
            else
            {
                // FIX (Ambiguity): Use the full class name for ParameterField.
                // The 'ParameterFields' collection uses the type from the 'Shared' namespace.
                foreach (CrystalDecisions.Shared.ParameterField param in reportDocument.ParameterFields)
                {
                    sb.AppendLine($"Name: {param.Name}, Type: {param.ParameterValueType}, Prompt: \"{param.PromptText}\"");
                }
            }
        }

        /// <summary>
        /// Extracts the Record and Group selection formulas.
        /// </summary>
        private static void ExtractSelectionFormulas(ReportDocument reportDocument, StringBuilder sb)
        {
            sb.AppendLine("\n--- SELECTION FORMULAS ---");
            sb.AppendLine($"[Record Selection]: {reportDocument.RecordSelectionFormula}");
            sb.AppendLine($"[Group Selection]: {reportDocument.DataDefinition.GroupSelectionFormula}");
        }

        /// <summary>
        /// Extracts grouping information and sort direction.
        /// This version contains the fix for the 'DataДeports' typo.
        /// </summary>
        private static void ExtractGrouping(ReportDocument reportDocument, StringBuilder sb)
        {
            sb.AppendLine("\n--- GROUPING ---");
            if (reportDocument.DataDefinition.Groups.Count == 0) { sb.AppendLine("No groups found."); return; }
            for (int i = 0; i < reportDocument.DataDefinition.Groups.Count; i++)
            {
                // FIX: Corrected 'DataДeports' to 'DataDefinition'.
                CrystalDecisions.CrystalReports.Engine.Group group = reportDocument.DataDefinition.Groups[i];

                SortField sortField = reportDocument.DataDefinition.SortFields.Cast<SortField>().FirstOrDefault(sf => sf.Field.Name == group.ConditionField.Name);
                string sortDirection = "(unknown)";
                if (sortField != null)
                {
                    sortDirection = sortField.SortDirection == SortDirection.AscendingOrder ? "Ascending" : "Descending";
                }
                sb.AppendLine($"Group #{i + 1}: By Field [{group.ConditionField.Name}], Sort: {sortDirection}");
            }
        }

        /// <summary>
        /// Extracts the text of all custom formulas in the report.
        /// </summary>
        private static void ExtractCustomFormulas(ReportDocument reportDocument, StringBuilder sb)
        {
            sb.AppendLine("\n--- CUSTOM FORMULAS ---");
            if (reportDocument.DataDefinition.FormulaFields.Count == 0) { sb.AppendLine("No custom formulas found."); }
            else { foreach (FormulaFieldDefinition f in reportDocument.DataDefinition.FormulaFields) sb.AppendLine($"\n[Formula: {f.Name}]\n{f.Text}"); }
        }
    }
}