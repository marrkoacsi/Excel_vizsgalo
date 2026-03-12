using ClosedXML.Excel;
using System.Collections.Generic;

namespace Excel.Helpers
{
    public static class ExcelHelper
    {
        public static List<string> ReadHeaderColumns(string filePath)
        {
            var result = new List<string>();

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);
                int lastColumn = ws.LastColumnUsed().ColumnNumber();

                for (int i = 1; i <= lastColumn; i++)
                    result.Add(XLHelper.GetColumnLetterFromNumber(i));
            }

            return result;
        }
    }
}