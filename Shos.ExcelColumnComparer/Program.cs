using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Shos.ExcelTextReplacer
{
    static class Helper
    {
        public static bool IsBetween(this int @this, int minimum, int maximum) => @this >= minimum && @this <= maximum;
    }

    static class CsvHelper
    {
        const char comma           = ',';
        const char doubleQuoration = '\"';
        const char newLine         = '\n';
        const char carriageReturn  = '\r';

        public static char Separator { get; set; } = comma;

        public static string ToCsv(this IEnumerable<string> @this)
        {
            var stringBuilder = new StringBuilder();
            var count = 0;
            foreach (var text in @this) {
                if (count++ != 0)
                    stringBuilder.Append(Separator);
                stringBuilder.Append(text.ToCsv());
            }
            return stringBuilder.ToString();
        }

        static string ToCsv(this string @this)
        {
            var csv = @this.Replace(new string(doubleQuoration, 1), new string(doubleQuoration, 2));
            return csv.NeedsDoubleQuorations() ? doubleQuoration + csv + doubleQuoration : csv;
        }

        static bool NeedsDoubleQuorations(this string @this)
            => @this.Any(character => character == Separator || character == comma || character == doubleQuoration || character == newLine || character == carriageReturn);
    }

    static class Program
    {
        //const int maxRow = 100;

        static bool FullRowEnabled { get; set; } = false;

        static void Main(string[] args)
        {
            var parameters = GetParameters(args);

            if (parameters == null)
                Usage();
            else
                Compare(parameters.Value.targetExcelFilePath, parameters.Value.column1, parameters.Value.column2);
        }

        static (string targetExcelFilePath, int column1, int column2)? GetParameters(string[] args)
        {
            string? targetFilePath = null;
            int? column1 = null;
            int? column2 = null;
            for (var index = 0; index < args.Length; index++) {
                switch (args[index]) {
                    case "-i": case "-I": case "/i": case "/I":
                        if (args.Length > index + 1) {
                            var filePath = ToFullPath(args[index + 1]);
                            if (!string.IsNullOrWhiteSpace(filePath))
                                targetFilePath = filePath;
                            index++;
                        }
                        break;

                    case "-c": case "-C": case "/c": case "/C":
                        if (args.Length > index + 1) {
                            var columnsText = args[index + 1];
                            var columnTexts = columnsText.Split(',');
                            if (columnTexts.Length >= 2 && int.TryParse(columnTexts[0], out var number1) && int.TryParse(columnTexts[1], out var number2))
                                (column1, column2) = (number1, number2);
                            index++;
                        }
                        break;

                    case "-f": case "-F": case "/f": case "/F":
                        FullRowEnabled = true;
                        break;
                }
            }
            if (targetFilePath is null || column1 is null || column2 is null)
                return null;
            return (targetFilePath, column1.Value, column2.Value);
        }

        static string? ToFullPath(string filePath)
            => File.Exists(filePath) ? Path.GetFullPath(filePath) : null;

        static void Compare(string targetExcelFilePath, int column1, int column2)
        {
            var excel = new Excel.Application();
            excel.Visible = true;
            Compare(excel, targetExcelFilePath, column1, column2);
            excel.Quit();
        }

        static void Compare(Excel.Application excel, string targetExcelFilePath, int column1, int column2)
        {
            var workbook = excel.Workbooks.Open(targetExcelFilePath);
            Compare(workbook, column1, column2);
            workbook.Close(true);
        }

        static void Compare(Excel.Workbook workbook, int column1, int column2)
        {
            Console.WriteLine($"column1: {column1}, column2: {column2}");

            var sheetCount      = 0;
            var differenceCount = 0;
            foreach (Excel.Worksheet sheet in workbook.Sheets)
                Compare(column1, column2, ++sheetCount, ref differenceCount, sheet);
            Console.WriteLine($"Difference count: {differenceCount}");
        }

        static void Compare(int column1, int column2, int sheetCount, ref int differenceCount, Excel.Worksheet sheet)
        {
            //sheet.Select();
            var columnCount          = sheet.UsedRange.Columns.Count;
            var rowCount             = sheet.UsedRange.Rows   .Count;
            Console.WriteLine($"sheet{sheetCount} - row count: {rowCount}, column count: {columnCount}");
            var sheetDifferenceCount = 0;

            if (column1.IsBetween(1, columnCount) && column2.IsBetween(1, columnCount)) {
                for (var row = 1; row <= rowCount; row++) {
                    var text1 = ToText(sheet.Cells[row, column1]);
                    var text2 = ToText(sheet.Cells[row, column2]);
                    if (!text1.Equals(text2)) {
                        var text = FullRowEnabled
                                   ? ToTexts(sheet, row, columnCount).ToCsv()
                                   : $"(row:{row}, column1:{column1}) - [{text1}] - (row:{row}, column2:{column2}) - [{text2}]";
                        Console.WriteLine(text);
                        sheetDifferenceCount++;
                        differenceCount++;
                    }
                }
            }
            Console.WriteLine($"Sheet difference count: {sheetDifferenceCount}");
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

        static IEnumerable<string> ToTexts(Excel.Worksheet sheet, int row, int columnCount)
        {
            for (var column = 1; column <= columnCount; column++)
                yield return ToText(sheet.Cells[row, column]);
        }

        static void Usage() => Console.WriteLine(
            "Usage:\nShos.ExcelColumnComparer -i targetExcelFilePath -c column1,column2 -f\n" +
            "\n" +
            "-i targetExcelFilePath\tExcel file path.\n" +
            "-c column1,column2\tFirst column and second column to compare.\n" +
            "-f\t\t\tShow full row.\n" +
            "\n" +
            "ex.\n" +
            "\n" +
            "Shos.ExcelColumnComparer -i xxx.xlsx -c 1,2"
        );
    }
}
