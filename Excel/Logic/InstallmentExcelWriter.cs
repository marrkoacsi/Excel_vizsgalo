using ClosedXML.Excel;
using Excel.Models;
using System.Collections.Generic;

namespace Excel.Logic
{
    public class InstallmentExcelWriter
    {

        public void Write(string path, List<InstallmentPreviewRow> rows)
        {

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Részletfizetés");

            ws.Cell(1, 1).Value = "Lakásszám";
            ws.Cell(1, 2).Value = "Név";
            ws.Cell(1, 3).Value = "Bruttó";
            ws.Cell(1, 4).Value = "Nettó";
            ws.Cell(1, 5).Value = "ÁFA";
            ws.Cell(1, 6).Value = "Devizás";
            ws.Cell(1, 7).Value = "Számlaszám";

            int r = 2;

            foreach (var row in rows)
            {

                ws.Cell(r, 1).Value = row.Apartment;
                ws.Cell(r, 2).Value = row.Name;
                ws.Cell(r, 3).Value = row.Brutto;
                ws.Cell(r, 3).Style.NumberFormat.Format = "#,##0";
                ws.Cell(r, 4).Value = row.Netto;
                ws.Cell(r, 4).Style.NumberFormat.Format = "#,##0";
                ws.Cell(r, 5).Value = row.Afa;
                ws.Cell(r, 6).Value = row.Devizas;
                ws.Cell(r, 7).Value = row.InvoiceNumber;

                r++;
            }

            ws.Columns().AdjustToContents();

            wb.SaveAs(path);
        }
    }
}