using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Shos.ExcelTextReplacer
{
    class Program
    {
        const int maxRow = 100;

        static void Main(string[] args)
        {
            (string targetExcelFilePath, string replacementListExcelFilePath)? paths = GetFilePaths(args);

            if (paths == null)
                Usage();
            else
                Replace(paths.Value.targetExcelFilePath, paths.Value.replacementListExcelFilePath);
        }

        static (string, string)? GetFilePaths(string[] args)
        {
            if (args.Length >= 2) {
                var targetExcelFilePath          = args[0];
                var replacementListExcelFilePath = args[1];
                if (File.Exists(targetExcelFilePath) && File.Exists(replacementListExcelFilePath))
                    return (targetExcelFilePath, replacementListExcelFilePath);
            }
            return null;
        }

        static void Replace(string targetExcelFilePath, string replacementListExcelFilePath)
        {
            var excel = new Excel.Application();
            excel.Visible = true;
            var replacementTable = CreateReplacementTable(excel, replacementListExcelFilePath);
            var replacementCount = Replace(excel, targetExcelFilePath, replacementListExcelFilePath, replacementTable);
            Console.WriteLine($"replacement count: {replacementCount}");
            excel.Quit();
        }

        static string ToText(object cell)
        {
            var range = cell as Excel.Range;
            if (range != null) {
                dynamic value = range.Value;
                var text = Convert.ToString(value);
                if (!string.IsNullOrWhiteSpace(text))
                    return text;
            }
            return "";
        }
         
        static Dictionary<string, string> CreateReplacementTable(Excel.Application excel, string replacementListExcelFilePath)
        {
            var workbook         = excel.Workbooks.Open(replacementListExcelFilePath);
            var replacementTable = new Dictionary<string, string>();

            foreach (Excel.Worksheet sheet in workbook.Sheets) {
                for (int row = 1; row < /*short.MaxValue*/ maxRow; row++) {
                    var fromText = ToText(sheet.Cells[row, 1]);
                    var toText   = ToText(sheet.Cells[row, 2]);
                    if (!string.IsNullOrWhiteSpace(fromText) && !string.IsNullOrWhiteSpace(toText)) {
                        replacementTable[fromText] = toText;
                        Console.WriteLine($"ReplacementTable(row: {row}) {fromText} → {toText}");
                    } 
                }
            }

            workbook.Close(false);
            return replacementTable;
        }

        static int Replace(Excel.Application excel, string targetExcelFilePath, string replacementListExcelFilePath, Dictionary<string, string> replacementTable)
        {
            var workbook = excel.Workbooks.Open(targetExcelFilePath);
            var replacementCount = Replace(workbook, replacementTable);
            workbook.Close(true);
            return replacementCount;
        }

        static int Replace(Excel.Workbook workbook, Dictionary<string, string> replacementTable)
        {
            var sheetCount       = 0;
            var replacementCount = 0;
            foreach (Excel.Worksheet sheet in workbook.Sheets) {
                sheetCount++;
                //sheet.Select();

                var rowCount    = sheet.UsedRange.Rows   .Count;
                var columnCount = sheet.UsedRange.Columns.Count;
                Console.WriteLine($"sheet{sheetCount} - row count: {rowCount}, column count: {columnCount}");

                for (int row = 1; row < rowCount; row++) {
                    for (int column = 1; column < columnCount; column++) {
                        var range = sheet.Cells[row, column] as Excel.Range;
                        if (range != null) {
                            dynamic val = range.Value;
                            var text = Convert.ToString(val);
                            var toText = "";
                            if (!string.IsNullOrWhiteSpace(text) && replacementTable.TryGetValue(text, out toText)) {
                                range.Value = toText;
                                replacementCount++;
                                Console.WriteLine($"{replacementCount} Replace(row:{row},column:{column}) {text} → {toText}");
                            }
                        }
                    }
                }
            }
            return replacementCount;
        }

        static void Usage() => Console.WriteLine(
            "Usage:\nShos.ExcelTextReplacer [targetExcelFilePath] [replacementListExcelFilePath]\n" +
            "\n" +
            "ex.\n" +
            "\n" +
            "- targetExcelFile (before):\n" +
            "\n" +
            "\t\tapple,iphone,3\n" +
            "\t\tapple,IPHONE,2\n" +
            "\t\tYahoo,PIXEL,1\n" +
            "\n" +
            "- replacementListExcelFile:\n" +
            "\n" +
            "\t\told text,new text\n" +
            "\t\tapple,Apple\n" +
            "\t\tIPHONE,iPhone\n" +
            "\t\tiphone,iPhone\n" +
            "\t\tYahoo,Google\n" +
            "\n" +
            "- targetExcelFile (after):\n" +
            "\n" +
            "\t\tApple,iPhone,3\n" +
            "\t\tApple,iPhone,2\n" +
            "\t\tGoogle,PIXEL,1\n"
        );
    }
}
