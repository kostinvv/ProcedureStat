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
                    group => new List<string>
                    {
                        RequiresConversion(group) ? "Нет" : "Да",
                        CanConvert(group) ? "Да" : "Есть ошибки",
                        group.First()[Constant.ObjectFamilyCount].ToString() == "3" ? "3" 
                            : $"(Обратить внимание) {group.First()[Constant.ObjectFamilyCount].ToString()}",

                        group.First()[Constant.ObjectFamilyMinComplexity].ToString(),
                        group.First()[Constant.ObjectFamilyMaxComplexity].ToString()
                    }
                );

            var spreadsheetProcessor = new SpreadsheetProcessor(inputFilePath, outputFilePath);
            spreadsheetProcessor.ProcessDocument(procedureDetails, _columns);
        }

        private static bool RequiresConversion(IGrouping<string, DataRow> group) 
            => group.All(row => row.Field<string>(Constant.ObjectFamilyMaxComplexity) == "1");
        
        private static bool CanConvert(IGrouping<string, DataRow> group) 
            => group.All(row => !row.Field<string>(Constant.ObjectFamilyMinComplexity).StartsWith("-") && 
            !row.Field<string>(Constant.ObjectFamilyMaxComplexity).StartsWith("-"));
    }
}
