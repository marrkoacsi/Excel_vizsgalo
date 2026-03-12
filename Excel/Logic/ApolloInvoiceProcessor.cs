using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace Excel.Logic
{

    public class ApolloInvoiceProcessor
    {

        public async Task<List<(ApolloInvoice invoice, RealEstateInfo info)>> Process(string file)
        {
            

            var wb = new XLWorkbook(file);
            var ws = wb.Worksheet(1);

            var invoices = new List<ApolloInvoice>();

            string currentInvoice = null;

            string currentName = "";
            bool skipBlock = false;

            foreach (var row in ws.RowsUsed().Skip(1))
            {
                var nameCell = row.Cell("A").GetString();

                // ha új név kezdődik
                if (!string.IsNullOrWhiteSpace(nameCell))
                {
                    currentName = nameCell;
                }

                if (skipBlock)
                    continue;


                var invoiceCell = row.Cell("C").GetString();

                if (!string.IsNullOrWhiteSpace(invoiceCell))
                    currentInvoice = invoiceCell.Trim();

                double nettoUnit = ParseDouble(row.Cell("U").GetValue<string>());
                double qty = ParseDouble(row.Cell("R").GetValue<string>());
                double vat = ParseDouble(row.Cell("V").GetValue<string>());

                double netto = nettoUnit * qty;
                double brutto = netto * (1 + vat / 100);

                string description = row.Cell("Q").GetString();

                if (description.IndexOf("Lakáskiviteli többletköltség", StringComparison.OrdinalIgnoreCase) >= 0)
                    continue;

                var invoice = new ApolloInvoice
                {
                    Name = currentName,
                    InvoiceNumber = currentInvoice,
                    Description = description,
                    Netto = netto,
                    Brutto = brutto,
                    DocNumber = row.Cell("D").GetString()
                };

                invoices.Add(invoice);
            }

            var filtered = RemoveCancelled(invoices);

            filtered = RemoveOwnerChanges(filtered);

            var grouped = new Dictionary<string, (ApolloInvoice invoice, RealEstateInfo info)>();
            var seen = new HashSet<string>();

            foreach (var inv in filtered)
            {
                if (inv.Brutto <= 0)
                    continue;

                var info = ParseDescription(inv.Description);

                if (string.IsNullOrEmpty(info.Unit))
                    continue;

                // duplikátum kulcs
                string uniqueKey = inv.Name + "_" + inv.InvoiceNumber + "_" + info.Unit + "_" + inv.Brutto;

                if (seen.Contains(uniqueKey))
                    continue;

                seen.Add(uniqueKey);

                // csoportosítás számla + lakás szerint
                string groupKey = inv.InvoiceNumber + "_" + info.Unit;

                if (!grouped.ContainsKey(groupKey))
                {
                    grouped[groupKey] = (new ApolloInvoice
                    {
                        Name = inv.Name,
                        InvoiceNumber = inv.InvoiceNumber,
                        Description = inv.Description,
                        Netto = inv.Netto,
                        Brutto = inv.Brutto
                    }, info);
                }
                else
                {
                    var existing = grouped[groupKey];

                    existing.invoice.Brutto += inv.Brutto;
                    existing.invoice.Netto += inv.Netto;

                    grouped[groupKey] = existing;
                }
            }

            return grouped.Values.ToList();
        }


        private RealEstateInfo ParseDescription(string text)
        {
            var info = new RealEstateInfo();

            if (string.IsNullOrWhiteSpace(text))
                return info;

            text = text.ToUpper();

            // garázsban lévő tároló
            var storageGarage = Regex.Match(text, @"P\d-T\d{1,3}");
            if (storageGarage.Success)
            {
                info.Unit = storageGarage.Value;
                info.Type = "storage";
                info.Vat = 27;
                return info;
            }

            // garázs
            var garage = Regex.Match(text, @"P\d-\d{2,3}");
            if (garage.Success)
            {
                info.Unit = garage.Value;
                info.Type = "garage";
                info.Vat = 27;
                return info;
            }

            // külön tároló
            var storage = Regex.Match(text, @"T-\d{1,3}");
            if (storage.Success)
            {
                info.Unit = storage.Value;
                info.Type = "storage";
                info.Vat = 27;
                return info;
            }

            // lakás
            var apartment = Regex.Match(text, @"[A-Z]-\d{2,3}");
            if (apartment.Success)
            {
                info.Unit = apartment.Value;
                info.Type = "apartment";
                info.Vat = 5;
                return info;
            }

            info.Type = "unknown";
            info.Vat = 0;

            return info;
        }

        private List<ApolloInvoice> RemoveOwnerChanges(List<ApolloInvoice> invoices)
        {
            var result = new List<ApolloInvoice>();

            var grouped = invoices.GroupBy(x => ParseDescription(x.Description).Unit);

            foreach (var unitGroup in grouped)
            {
                if (string.IsNullOrEmpty(unitGroup.Key))
                {
                    result.AddRange(unitGroup);
                    continue;
                }

                var ownerTotals = unitGroup
                    .GroupBy(x => x.Name)
                    .Select(g => new
                    {
                        Name = g.Key,
                        Total = g.Sum(x => x.Brutto)
                    })
                    .ToList();

                if (ownerTotals.Count == 1)
                {
                    result.AddRange(unitGroup);
                    continue;
                }

                // az aktuális tulaj = legnagyobb összeg
                var currentOwner = ownerTotals
                    .OrderByDescending(x => x.Total)
                    .First().Name;

                foreach (var inv in unitGroup)
                {
                    if (inv.Name == currentOwner)
                        result.Add(inv);
                }
            }

            return result;
        }

        private List<ApolloInvoice> RemoveCancelled(List<ApolloInvoice> invoices)
        {
            var cancelledNumbers = new HashSet<string>();

            // sztornó számlák kigyűjtése
            foreach (var inv in invoices)
            {
                if (!string.IsNullOrWhiteSpace(inv.DocNumber) && inv.DocNumber.StartsWith("#"))
                {
                    var num = inv.DocNumber.Substring(1).Trim();
                    cancelledNumbers.Add(num);
                }
            }

            var result = new List<ApolloInvoice>();

            foreach (var inv in invoices)
            {
                // sztornó sor
                if (!string.IsNullOrWhiteSpace(inv.DocNumber) && inv.DocNumber.StartsWith("#"))
                    continue;

                // sztornózott számla
                if (!string.IsNullOrWhiteSpace(inv.InvoiceNumber) &&
                    cancelledNumbers.Contains(inv.InvoiceNumber.Trim()))
                    continue;

                result.Add(inv);
            }

            return result;
        }


        private double ParseDouble(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return 0;

            value = value.Replace(" ", "")
                         .Replace("HUF", "")
                         .Replace(",", ".");


            double.TryParse(
                value,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out double result
            );

            return result;
        }
    }
}