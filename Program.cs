using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ProcedureStat
{
    public class Program
    {
        private static List<string> _columns = new()
        {
            "Требует конвертации",
            "Возможность конвертации",
            "Кол-во процедур",
            "Мин. сложность",
            "Макс. сложность",
            "Insert",
            "Update",
            "Delete"
        };

        private static void Main(string[] args)
        {
            string procedureDetailsFilePath = args[0];
            string inputFilePath = args[1];
            string outputFilePath = args[2];

            var spreadsheetConverter = new SpreadsheetConverter(procedureDetailsFilePath);

            var procedureDetails = spreadsheetConverter
                .ConvertToDataTable()
                .AsEnumerable()
                .Where(row => row.Field<string>(Constant.Scheme) == Constant.SchemeName)
                .GroupBy(row => row.Field<string>(Constant.ObjectKey))
                .ToDictionary(
                    group => group.Key,
                    group => 
                    {
                        var errorOperations = new List<string>
                        {
                            CheckComplexityLevel(group, "insert"),
                            CheckComplexityLevel(group, "update"),
                            CheckComplexityLevel(group, "delete")
                        }.Where(op => op != null).ToList();

                        return new List<string>
                        {
                            RequiresConversion(group) ? "Нет" : "Да",
                            errorOperations.Count == 0 ? "Да" : $"Есть ошибки ({string.Join(", ", errorOperations)})",
                            group.First()[Constant.ObjectFamilyCount].ToString() == "3" ? "3"
                                : $"(Обратить внимание) {group.First()[Constant.ObjectFamilyCount]}",
                            group.First()[Constant.ObjectFamilyMinComplexity].ToString(),
                            group.First()[Constant.ObjectFamilyMaxComplexity].ToString(),
                            CheckDbOperation(group, "insert"),
                            CheckDbOperation(group, "update"),
                            CheckDbOperation(group, "delete")
                        };
                    }
                );

            var spreadsheetProcessor = new SpreadsheetProcessor(inputFilePath, outputFilePath);
            spreadsheetProcessor.ProcessDocument(procedureDetails, _columns);
        }

        private static string CheckDbOperation(IEnumerable<DataRow> group, string operation) 
            => group.Any(row => row[Constant.DbKey].ToString().Contains(operation)) ? "+" : "-";

        private static bool RequiresConversion(IGrouping<string, DataRow> group) 
            => group.All(row => row.Field<string>(Constant.ObjectFamilyMaxComplexity) == "1");
        
        private static bool CanConvert(IGrouping<string, DataRow> group) 
            => group.All(row => !row.Field<string>(Constant.ObjectFamilyMinComplexity).StartsWith("-") && 
            !row.Field<string>(Constant.ObjectFamilyMaxComplexity).StartsWith("-"));

        private static string CheckComplexityLevel(IEnumerable<DataRow> group, string operation)
        {
            return group.Any(row => row[Constant.DbKey].ToString().Contains(operation) &&
                                    row.Field<string>(Constant.Complexity).StartsWith("-")) ? operation : null;
        }
    }
}
