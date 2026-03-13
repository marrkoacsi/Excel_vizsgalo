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

        private string apolloTextColumn;

        private List<ModificationPreview> lastResult;

        // ----------------------------------------
        // RÉSZLETFIZETÉS
        // ----------------------------------------

        private List<string> installmentFiles = new List<string>();
        private string devizaFile;

        private List<InstallmentPreviewRow> installmentPreview =
            new List<InstallmentPreviewRow>();

        private bool IsStandardInvoice(string inv)
        {
            return System.Text.RegularExpressions.Regex
                .IsMatch(inv ?? "", @"^\d+\/\d{4}$");
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        DebugConsole debugWindow;

        private void OpenDebugConsole_Click(object sender, RoutedEventArgs e)
        {
            if (debugWindow == null || !debugWindow.IsLoaded)
            {
                debugWindow = new DebugConsole();
                DebugLogger.Console = debugWindow;
            }

            debugWindow.Show();
            debugWindow.Activate();
        }

        // ----------------------------------------
        // FILE BROWSERS
        // ----------------------------------------

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

        private string SelectFile()
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx"
            };

            return dlg.ShowDialog() == true ? dlg.FileName : null;
        }

        // ----------------------------------------
        // COLUMN SETTINGS
        // ----------------------------------------

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

        // ----------------------------------------
        // ANALYZE
        // ----------------------------------------

        private void Analyze_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(osszesitoPath) ||
                string.IsNullOrEmpty(modositandoPath))
            {
                MessageBox.Show("Összesítő és módosítandó fájl kötelező.");
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

        // ----------------------------------------
        // SAVE
        // ----------------------------------------

        private void SaveChanges_Click(object sender, RoutedEventArgs e)
        {
            if (lastResult == null || lastResult.Count == 0)
            {
                MessageBox.Show("Nincs menthető módosítás. Előbb futtasd az Előnézetet.");
                return;
            }

            if (string.IsNullOrEmpty(modositandoPath))
            {
                MessageBox.Show("Nincs kiválasztva a módosítandó fájl.");
                return;
            }

            outputStartColumn = OutputStartColumnBox.Text?.ToUpper();

            if (string.IsNullOrEmpty(outputStartColumn))
            {
                MessageBox.Show("Adj meg kezdő oszlopot.");
                return;
            }

            // SaveFileDialog popup
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Kimeneti fájl mentése",
                Filter = "Excel Files|*.xlsx",
                FileName = "output_merge",
                DefaultExt = ".xlsx"
            };

            if (dlg.ShowDialog() != true)
                return;

            string newPath = dlg.FileName;

            System.IO.File.Copy(modositandoPath, newPath, true);

            var writer = new Excel.Helpers.ExcelWriter();
            writer.ApplyChanges(newPath, lastResult, outputStartColumn);

            // Útvonal kijelzés frissítése
            if (OutputFilePathText != null)
                OutputFilePathText.Text = newPath;

            if (StatsText != null)
                StatsText.Text = $"Mentés kész: {newPath}";

            MessageBox.Show("Mentés kész:\n" + newPath);
        }

        // =========================================
        // RÉSZLETFIZETÉS TAB
        // =========================================

        private void BrowseInstallment_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Multiselect = true,
                Title = "Részletfizetés fájlok kiválasztása"
            };

            if (dlg.ShowDialog() == true)
            {
                // installmentFiles-ba megy, NEM apolloPaths-ba
                installmentFiles = dlg.FileNames.ToList();

                if (InstallmentFilesText != null)
                    InstallmentFilesText.Text = installmentFiles.Count == 1
                        ? installmentFiles[0]
                        : $"{installmentFiles.Count} fájl kiválasztva";
            }
        }


        // --- Deviza Tallózás: EGY fájl ---
        private void BrowseDeviza_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Multiselect = false,
                Title = "Deviza fájl kiválasztása"
            };

            if (dlg.ShowDialog() == true)
            {
                devizaFile = dlg.FileName;      // a meglévő private string devizaFile mezőbe

                if (DevizaFileText != null)
                    DevizaFileText.Text = dlg.FileName;
            }
        }

        private void BrowseDevizaFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx"
            };

            if (dlg.ShowDialog() == true)
            {
                devizaFile = dlg.FileName;
                DevizaFileText.Text = System.IO.Path.GetFileName(devizaFile);
            }
        }

        // ----------------------------------------
        // PREVIEW
        // ----------------------------------------

        private async void PreviewInstallment_Click(object sender, RoutedEventArgs e)
        {
            if (installmentFiles == null || installmentFiles.Count == 0)
            {
                MessageBox.Show("Válassz részletfizetés fájlt.");
                return;
            }

            installmentPreview.Clear();

            var processor = new ApolloInvoiceProcessor();
            var tempList = new List<InstallmentPreviewRow>();

            // --- Részletfizetés fájlok ---
            foreach (var file in installmentFiles)
            {
                var result = await processor.Process(file);

                foreach (var item in result)
                {
                    tempList.Add(new InstallmentPreviewRow
                    {
                        Apartment = item.info.Unit,
                        Name = item.invoice.Name,
                        BlockName = "",
                        Brutto = item.invoice.Brutto,
                        Netto = item.invoice.Netto,
                        Afa = item.info.Vat,
                        Devizas = "",           // NEM devizás
                        InvoiceNumber = item.invoice.InvoiceNumber
                    });
                }
            }

            // --- Deviza fájl: ugyanaz a processor, "Devizás" megjelölés ---
            if (!string.IsNullOrEmpty(devizaFile) && System.IO.File.Exists(devizaFile))
            {
                var deviza = await processor.Process(devizaFile);

                foreach (var item in deviza)
                {
                    tempList.Add(new InstallmentPreviewRow
                    {
                        Apartment = item.info.Unit,
                        Name = item.invoice.Name,
                        BlockName = "",
                        Brutto = item.invoice.Brutto,
                        Netto = item.invoice.Netto,
                        Afa = item.info.Vat,
                        Devizas = "Devizás",    // MEGJELÖLVE
                        InvoiceNumber = item.invoice.InvoiceNumber
                    });
                }
            }

            // --- Duplikátum szűrés ---
            installmentPreview = tempList
                .GroupBy(x => new { x.Name, x.InvoiceNumber, x.Apartment, x.Brutto })
                .Select(g => g.First())
                .ToList();

            // --- Grid frissítés ---
            InstallmentPreviewGrid.ItemsSource = null;
            InstallmentPreviewGrid.ItemsSource = installmentPreview;

            var ci = new System.Globalization.CultureInfo("hu-HU");
            int devizaCount = installmentPreview.Count(x => x.Devizas == "Devizás");

            if (ApolloStatsText != null)
                ApolloStatsText.Text = $"{installmentPreview.Count} tétel  |  {devizaCount} devizás";

            if (StatsText != null)
                StatsText.Text = $"Előnézet kész: {installmentPreview.Count} tétel ({devizaCount} devizás)";
        }

        // ----------------------------------------
        // SAVE
        // ----------------------------------------

        private void SaveInstallment_Click(object sender, RoutedEventArgs e)
        {
            if (installmentPreview == null || installmentPreview.Count == 0)
            {
                MessageBox.Show("Nincs menthető adat. Előbb futtasd az Előnézetet.");
                return;
            }

            // SaveFileDialog popup — fájlnév és útvonal megadása
            var dlg = new SaveFileDialog
            {
                Title = "Részletfizetés exportálása",
                Filter = "Excel Files|*.xlsx",
                FileName = "reszletfizetes_export",
                DefaultExt = ".xlsx"
            };

            if (dlg.ShowDialog() != true)
                return;

            // Duplikátum szűrés mentés előtt
            var filtered = installmentPreview
                .GroupBy(x => new { x.Name, x.InvoiceNumber, x.Apartment, x.Brutto })
                .Select(g => g.First())
                .ToList();

            var writer = new InstallmentExcelWriter();
            writer.Write(dlg.FileName, filtered);

            if (StatsText != null)
                StatsText.Text = $"Mentés kész: {dlg.FileName}";

            MessageBox.Show("Mentés kész:\n" + dlg.FileName);
        }

        // ================================================================
        // EZEKET ADD HOZZÁ A MainWindow.xaml.cs VÉGÉHEZ
        // (a legutolsó '}' ZÁRÓjel ELÉ)
        // ================================================================

        // ---------------------------------------------------
        // RÉSZLETFIZETÉS — Apollo AI gomb
        // ---------------------------------------------------

        private async void RunApolloAI_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (installmentFiles == null || installmentFiles.Count == 0)
                {
                    MessageBox.Show("Előbb válaszd ki a részletfizetés fájl(ok)at.");
                    return;
                }

                var processor = new ApolloInvoiceProcessor();

                // ApolloInvoice és RealEstateInfo globális névtérben van — NEM Excel.Models
                var result = new List<(ApolloInvoice invoice, RealEstateInfo info)>();

                // Részletfizetés fájlok
                foreach (var path in installmentFiles)
                {
                    var partial = await processor.Process(path);
                    result.AddRange(partial);
                }

                // Deviza fájl — ugyanaz a logika
                string devizaPath = DevizaFileText?.Text;
                if (!string.IsNullOrWhiteSpace(devizaPath) &&
                    devizaPath != "Nincs fájl kiválasztva" &&
                    System.IO.File.Exists(devizaPath))
                {
                    var devizaPartial = await processor.Process(devizaPath);
                    result.AddRange(devizaPartial);
                }

                // Duplikátum szűrés
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

                // Kiírás
                var writer = new ApolloResultWriter();
                string output = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "apollo_result.xlsx");
                writer.Write(output, result);

                // UI: formázott számok
                var ci = new System.Globalization.CultureInfo("hu-HU");

                if (InstallmentPreviewGrid != null)
                    InstallmentPreviewGrid.ItemsSource = result.Select(x => new
                    {
                        Név = x.invoice.Name,
                        Számlaszám = x.invoice.InvoiceNumber,
                        Ingatlan = x.info.Unit,
                        Típus = x.info.Type,
                        Nettó = ((decimal)x.invoice.Netto).ToString("N0", ci),
                        Bruttó = ((decimal)x.invoice.Brutto).ToString("N0", ci),
                        AFA = x.info.Vat + " %"
                    }).ToList();

                if (ApolloStatsText != null)
                    ApolloStatsText.Text = $"{result.Count} tétel feldolgozva";

                if (StatsText != null)
                    StatsText.Text = $"Kész: {result.Count} tétel → {output}";

                MessageBox.Show($"Feldolgozás kész!\nTételek: {result.Count}\nFájl: {output}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hiba: " + ex.Message);
            }
        }

    }

}