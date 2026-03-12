using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace Excel
{
    public partial class ColumnSettingsWindow : Window
    {
        public string SelectedNameColumn { get; private set; }
        public string SelectedApartmentColumn { get; private set; }
        public List<string> SelectedOptionColumns { get; private set; }

        public ColumnSettingsWindow(List<string> columns)
        {
            InitializeComponent();

            NameColumnComboBox.ItemsSource = columns;
            ApartmentColumnComboBox.ItemsSource = columns;
            OptionColumnsListBox.ItemsSource = columns;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SelectedNameColumn = NameColumnComboBox.SelectedItem?.ToString();
            SelectedApartmentColumn = ApartmentColumnComboBox.SelectedItem?.ToString();
            SelectedOptionColumns = OptionColumnsListBox
                                        .SelectedItems
                                        .Cast<string>()
                                        .ToList();

            DialogResult = true;
            Close();
        }
    }
}