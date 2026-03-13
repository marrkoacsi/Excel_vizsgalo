namespace Excel.Models
{
    public class ModificationPreview
    {
        public string Apartment { get; set; }
        public string Name { get; set; }
        public string BlockName { get; set; }

        public decimal Brutto { get; set; }
        public decimal Netto { get; set; }
        public decimal Afa { get; set; }

        public string MatchType { get; set; }

        public string BruttoFormatted
        {
            get
            {
                var ci = new System.Globalization.CultureInfo("hu-HU");
                return Brutto == 0 ? "—" : Brutto.ToString("N0", ci);
            }
        }
        public string NettoFormatted
        {
            get
            {
                var ci = new System.Globalization.CultureInfo("hu-HU");
                return Netto == 0 ? "—" : Netto.ToString("N0", ci);
            }
        }

        private static string FormatHuf(decimal value)
           {
               if (value == 0) return "—";
               // Ezres szóköz elválasztó, 0 tizedesjegy
               var ci = new System.Globalization.CultureInfo("hu-HU");
               return value.ToString("N0", ci);   // hu-HU locale: szóköz az ezreselválasztó


           }


    }
}