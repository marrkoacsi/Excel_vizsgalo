using ClosedXML.Excel;
using System.Collections.Generic;

public class ApolloResultWriter
{
    public void Write(string path,
        List<(ApolloInvoice invoice, RealEstateInfo info)> data)
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Result");

        ws.Cell(1, 1).Value = "Name";
        ws.Cell(1, 2).Value = "Invoice";
        ws.Cell(1, 3).Value = "Unit";
        ws.Cell(1, 4).Value = "Type";
        ws.Cell(1, 5).Value = "Netto";
        ws.Cell(1, 6).Value = "Brutto";
        ws.Cell(1, 7).Value = "VAT";

        int r = 2;

        foreach (var (invoice, info) in data)
        {
            ws.Cell(r, 1).Value = invoice.Name;
            ws.Cell(r, 2).Value = invoice.InvoiceNumber;
            ws.Cell(r, 3).Value = info.Unit;
            ws.Cell(r, 4).Value = info.Type;
            ws.Cell(r, 6).Value = invoice.Brutto;
            ws.Cell(r, 6).Style.NumberFormat.Format = "#,##0";

            ws.Cell(r, 5).Value = invoice.Netto;
            ws.Cell(r, 5).Style.NumberFormat.Format = "#,##0";
            ws.Cell(r, 7).Value = info.Vat;

            r++;
        }

        wb.SaveAs(path);
    }
}