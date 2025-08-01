// =================================================================================
//  Crystal Reports Definition Extractor - FINAL WORKING VERSION
//  Includes fix for the ParameterField casting error.
// =================================================================================
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace CrystalReportDocumenter_Workaround
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2) { Console.WriteLine("Usage: CrystalReportDocumenter.exe \"<in.rpt>\" \"<out.txt>\""); return; }
            string reportPath = args[0];
            string outputPath = args[1];
            if (!File.Exists(reportPath)) { Console.WriteLine($"Error: Input file not found."); return; }

            Console.WriteLine($"Processing: {Path.GetFileName(reportPath)}");
            var sb = new StringBuilder();
            ReportDocument reportDocument = new ReportDocument();
            try
            {
                reportDocument.Load(reportPath);

                sb.AppendLine($"# REPORT DEFINITION (Generated via Workaround): {Path.GetFileName(reportPath)}");
                sb.AppendLine("====================================================================");

                ExtractSqlQueryWithReflection(reportDocument, sb);
                ExtractParameters(reportDocument, sb); // This method is now fixed
                ExtractSelectionFormulas(reportDocument, sb);
                ExtractCustomFormulas(reportDocument, sb);
                ExtractGrouping(reportDocument, sb);

                File.WriteAllText(outputPath, sb.ToString());
                Console.WriteLine($"\nSuccessfully created definition file.");
            }
            catch (Exception ex) { Console.WriteLine($"\nAn error occurred: {ex.Message}\n{ex.StackTrace}"); }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();
            }
        }

        private static void ExtractSqlQueryWithReflection(ReportDocument rd, StringBuilder sb)
        {
            sb.AppendLine("\n--- DATABASE & SQL QUERY ---");
            try
            {
                if (!rd.IsLoaded) throw new ArgumentException("Report document is not loaded.");

                PropertyInfo pi = rd.Database.Tables.GetType().GetProperty("RasTables", BindingFlags.NonPublic | BindingFlags.Instance);
                if (pi != null)
                {
                    dynamic rasTables = pi.GetValue(rd.Database.Tables, null);
                    if (rasTables.Count > 0)
                    {
                        var commandText = rasTables[0].CommandText;
                        if (!string.IsNullOrEmpty(commandText))
                        {
                            sb.AppendLine("[SQL Source: Command Object (via Reflection)]");
                            sb.AppendLine(commandText);
                            return;
                        }
                    }
                }

                sb.AppendLine("[SQL Source: Generated from Linked Tables]");
                sb.AppendLine("Could not extract generated SQL because the standard API (GetSQLStatement) is inaccessible in this environment.");
                sb.AppendLine("The tables used are:");
                foreach (Table table in rd.Database.Tables) { sb.AppendLine($"- {table.Name}"); }
            }
            catch (Exception ex) { sb.AppendLine($"Could not retrieve SQL Query. Error: {ex.Message}"); }
        }

        // ========================================================
        // THIS IS THE CORRECTED METHOD
        // ========================================================
        private static void ExtractParameters(ReportDocument reportDocument, StringBuilder sb)
        {
            sb.AppendLine("\n--- PARAMETERS ---");
            if (reportDocument.ParameterFields.Count == 0)
            {
                sb.AppendLine("No parameters found.");
            }
            else
            {
                // FIX: The collection contains 'ParameterField' objects.
                foreach (ParameterField param in reportDocument.ParameterFields)
                {
                    // FIX: Use the properties available on the 'ParameterField' type.
                    sb.AppendLine($"Name: {param.Name}, Type: {param.ParameterValueType}, Prompt: \"{param.PromptText}\"");
                }
            }
        }

        private static void ExtractGrouping(ReportDocument reportDocument, StringBuilder sb)
        {
            sb.AppendLine("\n--- GROUPING ---");
            if (reportDocument.DataDefinition.Groups.Count == 0) { sb.AppendLine("No groups found."); return; }
            for (int i = 0; i < reportDocument.DataDefinition.Groups.Count; i++)
            {
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

        private static void ExtractSelectionFormulas(ReportDocument reportDocument, StringBuilder sb)
        {
            sb.AppendLine("\n--- SELECTION FORMULAS ---");
            sb.AppendLine($"[Record Selection]: {reportDocument.RecordSelectionFormula}");
            sb.AppendLine($"[Group Selection]: {reportDocument.DataDefinition.GroupSelectionFormula}");
        }

        private static void ExtractCustomFormulas(ReportDocument reportDocument, StringBuilder sb)
        {
            sb.AppendLine("\n--- CUSTOM FORMULAS ---");
            if (reportDocument.DataDefinition.FormulaFields.Count == 0) { sb.AppendLine("No custom formulas found."); }
            else { foreach (FormulaFieldDefinition f in reportDocument.DataDefinition.FormulaFields) sb.AppendLine($"\n[Formula: {f.Name}]\n{f.Text}"); }
        }
    }
}