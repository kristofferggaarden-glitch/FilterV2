using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using Ookii.Dialogs.Wpf;

namespace FilterV1
{
    public partial class ProcessRawDataWindow : Window
    {
        private readonly string _settingsFilePath;
        private RawDataSettings _settings;
        private List<string> _allRawFiles;
        private string _selectedRawFile;
        private Action<string> _onProcessComplete;

        public ProcessRawDataWindow(Action<string> onProcessComplete = null)
        {
            InitializeComponent();
            _onProcessComplete = onProcessComplete;

            // Setup settings file path
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appFolder = Path.Combine(appDataPath, "FilterV1");
            Directory.CreateDirectory(appFolder);
            _settingsFilePath = Path.Combine(appFolder, "RawDataSettings.json");

            _allRawFiles = new List<string>();
            LoadSettings();
            UpdateUI();
            LoadRawFiles();
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    string json = File.ReadAllText(_settingsFilePath);
                    _settings = JsonSerializer.Deserialize<RawDataSettings>(json) ?? new RawDataSettings();
                }
                else
                {
                    _settings = new RawDataSettings();
                    SaveSettings();
                }
            }
            catch (Exception)
            {
                _settings = new RawDataSettings();
            }
        }

        private void SaveSettings()
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(_settings, options);
                File.WriteAllText(_settingsFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Kunne ikke lagre innstillinger: {ex.Message}", "Feil",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void UpdateUI()
        {
            RawFileLocationTextBox.Text = string.IsNullOrEmpty(_settings.RawFileLocation)
                ? "Ikke valgt"
                : _settings.RawFileLocation;
        }

        private void LoadRawFiles()
        {
            _allRawFiles.Clear();
            RawFilesListBox.Items.Clear();

            if (string.IsNullOrEmpty(_settings.RawFileLocation) || !Directory.Exists(_settings.RawFileLocation))
            {
                StatusTextBlock.Text = "Råfil-lokasjon er ikke satt. Klikk 'Endre innstillinger'.";
                return;
            }

            try
            {
                var files = Directory.GetFiles(_settings.RawFileLocation, "*.xlsx")
                    .Concat(Directory.GetFiles(_settings.RawFileLocation, "*.xls"))
                    .Select(Path.GetFileName)
                    .OrderBy(f => f)
                    .ToList();

                _allRawFiles.AddRange(files);
                foreach (var file in files)
                {
                    RawFilesListBox.Items.Add(file);
                }

                StatusTextBlock.Text = $"Fant {files.Count} Excel-filer i råfil-mappen.";
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"Feil ved lasting av filer: {ex.Message}";
            }
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string searchTerm = SearchTextBox.Text?.Trim() ?? "";
            RawFilesListBox.Items.Clear();

            var filtered = string.IsNullOrEmpty(searchTerm)
                ? _allRawFiles
                : _allRawFiles.Where(f => f.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0);

            foreach (var file in filtered)
            {
                RawFilesListBox.Items.Add(file);
            }
        }

        private void RawFilesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (RawFilesListBox.SelectedItem != null)
            {
                _selectedRawFile = RawFilesListBox.SelectedItem.ToString();
            }
        }

        private void SelectTargetFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new VistaFolderBrowserDialog
            {
                Description = "Velg målmappe hvor filene skal kopieres",
                UseDescriptionForTitle = true
            };

            if (dialog.ShowDialog() == true)
            {
                TargetFolderTextBox.Text = dialog.SelectedPath;
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new RawDataSettingsWindow(_settings, updatedSettings =>
            {
                _settings = updatedSettings;
                SaveSettings();
                UpdateUI();
                LoadRawFiles();
            });
            settingsWindow.Owner = this;
            settingsWindow.ShowDialog();
        }

        private void ProcessButton_Click(object sender, RoutedEventArgs e)
        {
            // Validation
            if (string.IsNullOrEmpty(TargetFolderTextBox.Text) || !Directory.Exists(TargetFolderTextBox.Text))
            {
                MessageBox.Show("Vennligst velg en gyldig målmappe.", "Valideringsfeil",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string orderNumber = OrderNumberTextBox.Text?.Trim();
            if (string.IsNullOrEmpty(orderNumber))
            {
                MessageBox.Show("Vennligst angi et ordrenummer.", "Valideringsfeil",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(_selectedRawFile))
            {
                MessageBox.Show("Vennligst velg en råfil.", "Valideringsfeil",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(_settings.TemplateFile1) || !File.Exists(_settings.TemplateFile1) ||
                string.IsNullOrEmpty(_settings.TemplateFile2) || !File.Exists(_settings.TemplateFile2) ||
                string.IsNullOrEmpty(_settings.TemplateFile3) || !File.Exists(_settings.TemplateFile3))
            {
                MessageBox.Show("En eller flere malfiler er ikke konfigurert. Sjekk innstillinger.", "Valideringsfeil",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                ProcessRawData(TargetFolderTextBox.Text, orderNumber, _selectedRawFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Feil under prosessering: {ex.Message}", "Feil",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBlock.Text = $"FEIL: {ex.Message}";
            }
        }

        private void ProcessRawData(string targetFolder, string orderNumber, string rawFileName)
        {
            StatusTextBlock.Text = "Starter prosessering...\n";

            // Step 1: Copy template files with new names
            string targetFileF = Path.Combine(targetFolder, $"{orderNumber} F.xlsx");
            string targetFileN = Path.Combine(targetFolder, $"{orderNumber} N.xlsx");
            string targetFileD = Path.Combine(targetFolder, $"{orderNumber} D.xlsx");

            StatusTextBlock.Text += $"Kopierer malfiler...\n";

            File.Copy(_settings.TemplateFile1, targetFileF, true);
            File.Copy(_settings.TemplateFile2, targetFileN, true);
            File.Copy(_settings.TemplateFile3, targetFileD, true);

            StatusTextBlock.Text += $"✓ Kopiert {Path.GetFileName(_settings.TemplateFile1)} → {orderNumber} F.xlsx\n";
            StatusTextBlock.Text += $"✓ Kopiert {Path.GetFileName(_settings.TemplateFile2)} → {orderNumber} N.xlsx\n";
            StatusTextBlock.Text += $"✓ Kopiert {Path.GetFileName(_settings.TemplateFile3)} → {orderNumber} D.xlsx\n\n";

            // Step 2: Read raw data from all sheets (column C and K from row 3)
            string rawFilePath = Path.Combine(_settings.RawFileLocation, rawFileName);
            var columnCData = new List<string>();
            var columnKData = new List<string>();

            StatusTextBlock.Text += $"Leser rådata fra {rawFileName}...\n";

            using (var rawWorkbook = new XLWorkbook(rawFilePath))
            {
                int sheetCount = rawWorkbook.Worksheets.Count;
                StatusTextBlock.Text += $"Fant {sheetCount} sheet(s) i råfilen\n";

                foreach (var sheet in rawWorkbook.Worksheets)
                {
                    StatusTextBlock.Text += $"  Prosesserer sheet: {sheet.Name}\n";

                    // Read column C (index 3) from row 3 onwards
                    int row = 3;
                    while (true)
                    {
                        var cellC = sheet.Cell(row, 3);
                        string valueC = cellC.GetString().Trim();
                        if (string.IsNullOrEmpty(valueC))
                            break;
                        columnCData.Add(valueC);
                        row++;
                    }

                    // Read column K (index 11) from row 3 onwards
                    row = 3;
                    while (true)
                    {
                        var cellK = sheet.Cell(row, 11);
                        string valueK = cellK.GetString().Trim();
                        if (string.IsNullOrEmpty(valueK))
                            break;
                        columnKData.Add(valueK);
                        row++;
                    }
                }
            }

            StatusTextBlock.Text += $"✓ Lest {columnCData.Count} rader fra kolonne C\n";
            StatusTextBlock.Text += $"✓ Lest {columnKData.Count} rader fra kolonne K\n\n";

            // Step 3: Write data to target file (F) in columns E and F from row 2
            StatusTextBlock.Text += $"Skriver data til {orderNumber} F.xlsx...\n";

            using (var targetWorkbook = new XLWorkbook(targetFileF))
            {
                var targetSheet = targetWorkbook.Worksheets.First();

                // Write column C data to column E (5) starting at row 2
                for (int i = 0; i < columnCData.Count; i++)
                {
                    targetSheet.Cell(i + 2, 5).Value = columnCData[i];
                }

                // Write column K data to column F (6) starting at row 2
                for (int i = 0; i < columnKData.Count; i++)
                {
                    targetSheet.Cell(i + 2, 6).Value = columnKData[i];
                }

                targetWorkbook.Save();
            }

            StatusTextBlock.Text += $"✓ Data skrevet til {orderNumber} F.xlsx\n";
            StatusTextBlock.Text += $"  - {columnCData.Count} rader i kolonne E\n";
            StatusTextBlock.Text += $"  - {columnKData.Count} rader i kolonne F\n\n";
            StatusTextBlock.Text += "=================================\n";
            StatusTextBlock.Text += "PROSESSERING FULLFØRT!\n";
            StatusTextBlock.Text += "=================================\n";

            MessageBox.Show("Prosessering fullført! F-filen vil nå lastes inn i hovedvinduet.", "Suksess",
                MessageBoxButton.OK, MessageBoxImage.Information);

            // Call the callback with the F file path to load it in MainWindow
            _onProcessComplete?.Invoke(targetFileF);

            // Close this window
            Close();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    public class RawDataSettings
    {
        public string RawFileLocation { get; set; } = "";
        public string TemplateFile1 { get; set; } = ""; // Fil som får data (F)
        public string TemplateFile2 { get; set; } = ""; // 4. Nord Ledningsliste - mal (N)
        public string TemplateFile3 { get; set; } = ""; // 3. Durapart ledningsliste - mal (D)
    }
}