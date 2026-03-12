using Excel.Helpers;
using Excel.Models;
using Excel.Logic;

using Microsoft.Win32;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace Excel
{
    public partial class MainWindow : Window
    {
        private string osszesitoPath;
        List<string> apolloPaths;
        private string modositandoPath;

        private string nameColumn;
        private string apartmentColumn;
        private List<string> optionColumns = new List<string>();

        private string apolloBruttoColumn;
        private string apolloNettoColumn;
        private string apolloAfaColumn;

        private string outputStartColumn;
        private string outputFileName;

        private string apolloTextColumn;

        private List<ModificationPreview> lastResult;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BrowseOsszesito_Click(object sender, RoutedEventArgs e)
        {
            osszesitoPath = SelectFile();
        }

        private void BrowseApollo_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Multiselect = true
            };

            if (dlg.ShowDialog() == true)
                apolloPaths = dlg.FileNames.ToList();
        }

        private void BrowseModositando_Click(object sender, RoutedEventArgs e)
        {
            modositandoPath = SelectFile();
        }

        private void OpenColumnSettings_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(osszesitoPath))
            {
                MessageBox.Show("Előbb válaszd ki az Összesítő fájlt.");
                return;
            }

            var columns = ExcelHelper.ReadHeaderColumns(osszesitoPath);

            var window = new ColumnSettingsWindow(columns)
            {
                Owner = this
            };

            if (window.ShowDialog() == true)
            {
                nameColumn = window.SelectedNameColumn;
                apartmentColumn = window.SelectedApartmentColumn;
                optionColumns = window.SelectedOptionColumns;
            }
        }

        private void Analyze_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(osszesitoPath) ||
                string.IsNullOrEmpty(modositandoPath))
            {
                MessageBox.Show("Összesítő és módosítandó fájl kötelező.");
                return;
            }

            if (string.IsNullOrEmpty(nameColumn) ||
                string.IsNullOrEmpty(apartmentColumn) ||
                optionColumns == null ||
                optionColumns.Count == 0)
            {
                MessageBox.Show("Állítsd be az oszlopokat az Oszlop beállítások gombbal.");
                return;
            }

            var processor = new ExcelProcessor();

            lastResult = processor.Process(
                osszesitoPath,
                apolloPaths,
                nameColumn,
                apartmentColumn,
                optionColumns,
                apolloBruttoColumn,
                apolloNettoColumn,
                apolloAfaColumn
            );

            PreviewGrid.ItemsSource = lastResult;
            StatsText.Text = "Talált módosítások száma: " + lastResult.Count;
        }

        private void SaveChanges_Click(object sender, RoutedEventArgs e)
        {
            if (lastResult == null || lastResult.Count == 0)
            {
                MessageBox.Show("Nincs menthető módosítás.");
                return;
            }

            if (string.IsNullOrEmpty(modositandoPath))
            {
                MessageBox.Show("Nincs kiválasztva a módosítandó fájl.");
                return;
            }

            outputFileName = OutputFileNameBox.Text;
            outputStartColumn = OutputStartColumnBox.Text?.ToUpper();

            if (string.IsNullOrEmpty(outputFileName))
            {
                MessageBox.Show("Adj meg fájlnevet.");
                return;
            }

            if (string.IsNullOrEmpty(outputStartColumn))
            {
                MessageBox.Show("Adj meg kezdő oszlopot.");
                return;
            }

            string newPath = System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(modositandoPath),
                outputFileName + ".xlsx");

            System.IO.File.Copy(modositandoPath, newPath, true);

            var writer = new ExcelWriter();
            writer.ApplyChanges(newPath, lastResult, outputStartColumn);

            MessageBox.Show("Mentés kész: " + newPath);
        }

        private string SelectFile()
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx"
            };
            return dlg.ShowDialog() == true ? dlg.FileName : null;
        }

        private void OpenApolloColumnSettings_Click(object sender, RoutedEventArgs e)
        {
            if (apolloPaths == null || apolloPaths.Count == 0)
            {
                MessageBox.Show("Előbb válaszd ki az Apollo fájlt.");
                return;
            }

            var columns = ExcelHelper.ReadHeaderColumns(apolloPaths[0]);

            var window = new ApolloColumnSettingsWindow(columns)
            {
                Owner = this
            };

            if (window.ShowDialog() == true)
            {
                apolloTextColumn = window.SelectedTextColumn;
                apolloBruttoColumn = window.SelectedBruttoColumn;
                apolloNettoColumn = window.SelectedNettoColumn;
                apolloAfaColumn = window.SelectedAfaColumn;
            }
        }


        // ---------------------------------------------------
        // AI APOLLO FEATURE
        // ---------------------------------------------------

        private async void RunApolloAI_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                if (apolloPaths == null || apolloPaths.Count == 0)
                {
                    MessageBox.Show("Előbb válaszd ki az Apollo Excel fájlt.");
                    return;
                }

                string apiKey = "AIzaSyDS9ODMCVklPSrkP6wR-Z_R75ofIvBqrIE";

                if (string.IsNullOrEmpty(apiKey))
                {
                    MessageBox.Show("Hiányzik az OpenAI API kulcs az App.config fájlban.");
                    return;
                }


                var processor = new ApolloInvoiceProcessor();

                var result = new List<(ApolloInvoice invoice, RealEstateInfo info)>();

                foreach (var path in apolloPaths)
                {
                    var partial = await processor.Process(path);
                    result.AddRange(partial);
                }

                // duplikátum szűrés
                result = result
                    .GroupBy(x => new
                    {
                        x.invoice.Name,
                        x.invoice.InvoiceNumber,
                        x.info.Unit,
                        x.invoice.Brutto
                    })
                    .Select(g => g.First())
                    .ToList();

                var writer = new ApolloResultWriter();

                string output = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "apollo_ai_result.xlsx"
                );

                writer.Write(output, result);

                MessageBox.Show("Apollo feldolgozás kész:\n" + output);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}