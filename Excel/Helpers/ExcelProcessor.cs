using ClosedXML.Excel;
using Excel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Excel.Helpers
{
    public class ExcelProcessor
    {
        private class ApolloEntry
        {
            public decimal Brutto { get; set; }
            public decimal Netto { get; set; }
            public decimal Afa { get; set; }
        }

        private decimal SafeDecimal(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty())
                return 0;

            try
            {
                return cell.GetValue<decimal>();
            }
            catch
            {
                decimal parsed;
                if (decimal.TryParse(
                    cell.GetString()
                        .Replace("Ft", "")
                        .Replace(" ", "")
                        .Replace(",", "."),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out parsed))
                {
                    return parsed;
                }
            }

            return 0;
        }

        public List<ModificationPreview> Process(
            string cePath,
            List<string> apolloPaths,
            string nameColumn,
            string apartmentColumn,
            List<string> optionColumns,
            string apolloBruttoColumn,
            string apolloNettoColumn,
            string apolloAfaColumn)
        {
            var result = new List<ModificationPreview>();

            var ceWb = new XLWorkbook(cePath);
            var ceWs = ceWb.Worksheet(1);

            var apolloIndex = BuildCombinedApolloIndex(
                apolloPaths,
                apolloBruttoColumn,
                apolloNettoColumn,
                apolloAfaColumn);

            int headerRowNumber = ceWs.RowsUsed()
                .First(r => r.Cells().Any(c => c.GetString().Trim() == "Vevő"))
                .RowNumber();

            var headerRow = ceWs.Row(headerRowNumber);

            foreach (var row in ceWs.RowsUsed().Where(r => r.RowNumber() > headerRowNumber))
            {
                string name = row.Cell(nameColumn).GetString();
                string rawApartments = row.Cell(apartmentColumn).GetString();

                var apartments = rawApartments
                    .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(a => NormalizeApartment(a))
                    .Where(a => !string.IsNullOrEmpty(a))
                    .ToList();

                foreach (var columnLetter in optionColumns)
                {
                    int colIndex = XLHelper.GetColumnNumberFromLetter(columnLetter);
                    decimal brutto = SafeDecimal(row.Cell(colIndex));

                    if (brutto <= 0)
                        continue;

                    string blockName = headerRow.Cell(colIndex)
                        .GetString()
                        .Replace("(Bruttó Ft)", "")
                        .Replace("költség", "")
                        .Trim();

                    if (apartments.Count == 1)
                    {
                        decimal netto = Math.Round(brutto / 1.05m, 2);
                        decimal afa = brutto - netto;

                        result.Add(new ModificationPreview
                        {
                            Apartment = apartments[0],
                            Name = name,
                            BlockName = blockName,
                            Brutto = brutto,
                            Netto = netto,
                            Afa = 5,
                            MatchType = "CE_single"
                        });
                    }
                    else
                    {
                        if (apolloIndex != null && apolloIndex.ContainsKey(brutto))
                        {
                            var apolloEntry = apolloIndex[brutto]
                                .GroupBy(x => new { x.Brutto, x.Netto, x.Afa })
                                .Select(g => g.First())
                                .FirstOrDefault();

                            if (apolloEntry != null)
                            {
                                result.Add(new ModificationPreview
                                {
                                    Apartment = apartments.FirstOrDefault(),
                                    Name = name,
                                    BlockName = blockName,
                                    Brutto = brutto,
                                    Netto = apolloEntry.Netto,
                                    Afa = apolloEntry.Afa,
                                    MatchType = "Apollo_match"
                                });
                            }
                        }

                        else
                        {
                            result.Add(new ModificationPreview
                            {
                                Apartment = apartments.FirstOrDefault(),
                                Name = name,
                                BlockName = blockName,
                                Brutto = brutto,
                                Netto = Math.Round(brutto / 1.05m, 2),
                                Afa = 5,
                                MatchType = "Manual_review"
                            });
                        }
                    }
                }
            }

            ceWb.Dispose();
            return result;
        }

        private Dictionary<decimal, List<ApolloEntry>> BuildCombinedApolloIndex(
            List<string> apolloPaths,
            string bruttoColumn,
            string nettoColumn,
            string afaColumn)
        {
            var index = new Dictionary<decimal, List<ApolloEntry>>();

            if (apolloPaths == null)
                return index;

            foreach (var path in apolloPaths)
            {
                using (var wb = new XLWorkbook(path))
                {
                    var ws = wb.Worksheet(1);

                    foreach (var row in ws.RowsUsed().Skip(1))
                    {
                        decimal brutto = SafeDecimal(
                            row.Cell(XLHelper.GetColumnNumberFromLetter(bruttoColumn)));

                        var entry = new ApolloEntry
                        {
                            Brutto = brutto,
                            Netto = SafeDecimal(
                                row.Cell(XLHelper.GetColumnNumberFromLetter(nettoColumn))),
                            Afa = SafeDecimal(
                                row.Cell(XLHelper.GetColumnNumberFromLetter(afaColumn)))
                        };

                        if (!index.ContainsKey(brutto))
                            index[brutto] = new List<ApolloEntry>();

                        index[brutto].Add(entry);
                    }
                }
            }

            return index;
        }

        private string NormalizeApartment(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return null;

            var match = Regex.Match(input.ToUpper(), @"([A-Z])\-?(\d+)");
            if (!match.Success)
                return null;

            return match.Groups[1].Value + "-" + int.Parse(match.Groups[2].Value);
        }
    }
}