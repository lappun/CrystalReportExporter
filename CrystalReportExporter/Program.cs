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
            // Main method remains the same...
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
                ExtractSqlQueryWithReflection(reportDocument, sb); // Using Plan B
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

        // ========================================================
        // THIS IS THE 'PLAN B' DEEP REFLECTION METHOD
        // ========================================================
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
                        object tableObject = rasTables[0];
                        var commandTextProperty = tableObject.GetType().GetProperty("CommandText");
                        if (commandTextProperty != null)
                        {
                            var commandText = commandTextProperty.GetValue(tableObject, null)?.ToString();
                            if (!string.IsNullOrEmpty(commandText))
                            {
                                sb.AppendLine("[SQL Source: Command Object (via Deep Reflection)]");
                                sb.AppendLine(commandText);
                                return; // Found a command, so we are done.
                            }
                        }
                    }
                }

                // ========================================================
                // THIS IS THE NEW "GOOD ENOUGH" PART
                // ========================================================
                sb.AppendLine("[SQL Source: Report uses Linked Tables]");
                sb.AppendLine("The standard API to generate the SQL is inaccessible in this environment.");
                sb.AppendLine("Listing the tables used instead:");

                if (rd.Database.Tables.Count > 0)
                {
                    foreach (Table table in rd.Database.Tables)
                    {
                        sb.AppendLine($"- {table.Name}");
                    }
                }
                else
                {
                    sb.AppendLine("No tables found.");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Could not retrieve database info. Error: {ex.Message}");
            }
        }

        // --- Other methods remain unchanged ---

        private static void ExtractParameters(ReportDocument reportDocument, StringBuilder sb)
        {
            sb.AppendLine("\n--- PARAMETERS ---");
            if (reportDocument.ParameterFields.Count == 0) { sb.AppendLine("No parameters found."); }
            else
            {
                foreach (ParameterField param in reportDocument.ParameterFields)
                {
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