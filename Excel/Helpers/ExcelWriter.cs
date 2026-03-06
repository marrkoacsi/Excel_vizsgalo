using ClosedXML.Excel;
using Excel.Models;
using System.Collections.Generic;
using System.Linq;

namespace Excel.Helpers
{
    public class ExcelWriter
    {
        public void ApplyChanges(
            string filePath,
            List<ModificationPreview> changes,
            string startColumnLetter)
        {
            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);

                int startColIndex = XLHelper.GetColumnNumberFromLetter(startColumnLetter);

                var modNames = changes
                    .Select(c => c.BlockName)
                    .Distinct()
                    .ToList();

                // ===== 1️⃣ FEJLÉC =====
                int headerCol = startColIndex;

                foreach (var mod in modNames)
                {
                    ws.Cell(1, headerCol).Value = mod;
                    ws.Cell(1, headerCol + 1).Value = "Nettó érték";
                    ws.Cell(1, headerCol + 2).Value = "Bruttó érték";
                    ws.Cell(1, headerCol + 3).Value = "ÁFA";

                    ws.Range(1, headerCol, 1, headerCol + 3)
                      .Style.Font.Bold = true;

                    headerCol += 4;
                }

                // ===== 2️⃣ Lakásszám mapping (A oszlop) =====
                var rowMap = new Dictionary<string, IXLRow>();

                foreach (var row in ws.RowsUsed().Skip(1))
                {
                    var raw = row.Cell("A").GetString();
                    var normalized = NormalizeApartment(raw);

                    if (!string.IsNullOrEmpty(normalized))
                        rowMap[normalized] = row;
                }

                int maxWrittenRow = 0;

                // ===== 3️⃣ ADAT KIÍRÁS =====
                foreach (var change in changes)
                {
                    if (string.IsNullOrEmpty(change.Apartment))
                        continue;

                    if (!rowMap.ContainsKey(change.Apartment))
                        continue;

                    var targetRow = rowMap[change.Apartment];

                    int modIndex = modNames.IndexOf(change.BlockName);
                    int baseCol = startColIndex + (modIndex * 4);

                    var nettoCell = ws.Cell(targetRow.RowNumber(), baseCol + 1);
                    var bruttoCell = ws.Cell(targetRow.RowNumber(), baseCol + 2);
                    var afaCell = ws.Cell(targetRow.RowNumber(), baseCol + 3);

                    decimal brutto = change.Brutto;
                    decimal netto = change.Netto;
                    decimal afa = change.Afa;

                    // 🔥 Ha csak bruttó van → számolunk 5%-kal
                    if (brutto > 0 && netto == 0)
                    {
                        netto = System.Math.Round(brutto / 1.05m, 2);
                        afa = 5;
                    }

                    nettoCell.Value = netto;
                    bruttoCell.Value = brutto;
                    afaCell.Value = afa;

                    // 🔴 csak ha tényleg nincs adat
                    if (brutto == 0)
                        bruttoCell.Style.Fill.BackgroundColor = XLColor.LightPink;

                    if (netto == 0 && brutto == 0)
                        nettoCell.Style.Fill.BackgroundColor = XLColor.LightPink;

                    if (afa == 0 && brutto == 0)
                        afaCell.Style.Fill.BackgroundColor = XLColor.LightPink;

                    if (targetRow.RowNumber() > maxWrittenRow)
                        maxWrittenRow = targetRow.RowNumber();
                }

                // ===== 4️⃣ ÖSSZESEN SOR =====
                if (maxWrittenRow > 1)
                {
                    int summaryRow = maxWrittenRow + 1;

                    ws.Cell(summaryRow, startColIndex).Value = "ÖSSZESEN";
                    ws.Cell(summaryRow, startColIndex).Style.Font.Bold = true;

                    for (int i = 0; i < modNames.Count; i++)
                    {
                        int baseCol = startColIndex + (i * 4);

                        string nettoColLetter = XLHelper.GetColumnLetterFromNumber(baseCol + 1);
                        string bruttoColLetter = XLHelper.GetColumnLetterFromNumber(baseCol + 2);

                        ws.Cell(summaryRow, baseCol + 1)
                            .FormulaA1 = $"SUM({nettoColLetter}2:{nettoColLetter}{maxWrittenRow})";

                        ws.Cell(summaryRow, baseCol + 2)
                            .FormulaA1 = $"SUM({bruttoColLetter}2:{bruttoColLetter}{maxWrittenRow})";
                    }

                    ws.Row(summaryRow).Style.Font.Bold = true;
                }

                wb.Save();
            }
        }

        private string NormalizeApartment(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return null;

            var match = System.Text.RegularExpressions.Regex
                .Match(input.ToUpper(), @"([A-Z])\-?(\d+)");

            if (!match.Success)
                return null;

            return match.Groups[1].Value + "-" + int.Parse(match.Groups[2].Value);
        }
    }
}