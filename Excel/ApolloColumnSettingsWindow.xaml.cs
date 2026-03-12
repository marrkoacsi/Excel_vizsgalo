using System.Collections.Generic;
using System.Windows;

namespace Excel
{
    public partial class ApolloColumnSettingsWindow : Window
    {
        public string SelectedNameColumn { get; private set; }
        public string SelectedBruttoColumn { get; private set; }
        public string SelectedNettoColumn { get; private set; }
        public string SelectedAfaColumn { get; private set; }
        public string SelectedTextColumn { get; set; }
        public string TetelSzoveg { get; set; }
        public ApolloColumnSettingsWindow(List<string> columns)
        {
            InitializeComponent();

            NameCombo.ItemsSource = columns;
            BruttoCombo.ItemsSource = columns;
            NettoCombo.ItemsSource = columns;
            AfaCombo.ItemsSource = columns;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SelectedNameColumn = NameCombo.SelectedItem?.ToString();
            SelectedBruttoColumn = BruttoCombo.SelectedItem?.ToString();
            SelectedNettoColumn = NettoCombo.SelectedItem?.ToString();
            SelectedAfaColumn = AfaCombo.SelectedItem?.ToString();

            DialogResult = true;
            Close();
        }
    }
}