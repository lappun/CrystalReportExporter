// =================================================================================
//  Crystal Reports Definition Extractor - WORKAROUND VERSION
//  Uses Reflection to bypass broken API access.
// =================================================================================
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared; // You still need this basic reference

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
                // We will try loading without the enum, as it's failing for you.
                reportDocument.Load(reportPath);

                sb.AppendLine($"# REPORT DEFINITION (Generated via Workaround): {Path.GetFileName(reportPath)}");
                sb.AppendLine("====================================================================");

                ExtractSqlQueryWithReflection(reportDocument, sb);
                ExtractParameters(reportDocument, sb);
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

        /// <summary>
        /// This method uses Reflection to get the SQL Command text. It's a workaround for
        /// when the standard API ('GetSQLStatement') fails due to environment issues.
        /// </summary>
        private static void ExtractSqlQueryWithReflection(ReportDocument rd, StringBuilder sb)
        {
            sb.AppendLine("\n--- DATABASE & SQL QUERY ---");
            try
            {
                if (!rd.IsLoaded)
                    throw new ArgumentException("Report document is not loaded.");

                // This works for reports based on a "SQL Command".
                // It uses a non-public property, so it's fragile and may break with future CR versions.
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
                            return; // Success, we are done.
                        }
                    }
                }

                // If the above fails, it's likely a report with linked tables.
                // Since GetSQLStatement is broken in your environment, we can't get the generated SQL.
                sb.AppendLine("[SQL Source: Generated from Linked Tables]");
                sb.AppendLine("Could not extract generated SQL because the standard API (GetSQLStatement) is inaccessible in this environment.");
                sb.AppendLine("The tables used are:");
                foreach (Table table in rd.Database.Tables)
                {
                    sb.AppendLine($"- {table.Name}");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Could not retrieve SQL Query. Error: {ex.Message}");
            }
        }

        private static void ExtractGrouping(ReportDocument reportDocument, StringBuilder sb)
        {
            sb.AppendLine("\n--- GROUPING ---");
            if (reportDocument.DataDefinition.Groups.Count == 0) { sb.AppendLine("No groups found."); return; }

            for (int i = 0; i < reportDocument.DataDefinition.Groups.Count; i++)
            {
                // This logic is correct and uses the basic API.
                CrystalDecisions.CrystalReports.Engine.Group group = reportDocument.DataDefinition.Groups[i];

                // FIX for 'SortDirection' not found:
                // Find the corresponding sort field to get the direction.
                SortField sortField = reportDocument.DataDefinition.SortFields
                    .Cast<SortField>()
                    .FirstOrDefault(sf => sf.Field.Name == group.ConditionField.Name);

                string sortDirection = "(unknown)";
                if (sortField != null)
                {
                    // The SortDirection enum is in CrystalDecisions.Shared.dll
                    sortDirection = sortField.SortDirection == SortDirection.AscendingOrder ? "Ascending" : "Descending";
                }

                sb.AppendLine($"Group #{i + 1}: By Field [{group.ConditionField.Name}], Sort: {sortDirection}");
            }
        }

        // --- These methods use the basic API and should work if your core references are okay ---
        private static void ExtractParameters(ReportDocument reportDocument, StringBuilder sb)
        {
            sb.AppendLine("\n--- PARAMETERS ---");
            if (reportDocument.ParameterFields.Count == 0) { sb.AppendLine("No parameters found."); }
            else { foreach (ParameterFieldDefinition p in reportDocument.ParameterFields) sb.AppendLine($"Name: {p.Name}, Type: {p.ValueType}"); }
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