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
    }
}