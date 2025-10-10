using ClosedXML.Excel;
using ExcelDataReader;
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

            // Registrer encoding provider for ExcelDataReader (påkrevd for .xls-filer)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Setup settings file path
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appFolder = Path.Combine(appDataPath, "FilterV1");
            Directory.CreateDirectory(appFolder);
            _settingsFilePath = Path.Combine(appFolder, "RawDataSettings.json");

            _allRawFiles = new List<string>();
            LoadSettings();
            UpdateUI();

            // When the order number changes, automatically update the search box with the base part
            // (the portion before any dash) so the user doesn't have to retype it under step 3.  This
            // event is registered after controls are initialized to avoid null references.
            OrderNumberTextBox.TextChanged += OrderNumberTextBox_TextChanged;

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

            // Populate the target folder text box from the last used folder if available.  When
            // nothing has been saved the field will remain blank, allowing the user to select a
            // destination folder manually.  This persists across sessions so repeated runs do not
            // require re‑selecting the same path.
            if (!string.IsNullOrWhiteSpace(_settings.LastTargetFolder) && Directory.Exists(_settings.LastTargetFolder))
            {
                TargetFolderTextBox.Text = _settings.LastTargetFolder;
            }
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
                // Persist the selected target folder so it is remembered between sessions.  This
                // simplifies subsequent runs by automatically populating the path when the window
                // is opened.  Save the settings immediately to avoid losing the value if the
                // application closes unexpectedly.
                _settings.LastTargetFolder = dialog.SelectedPath;
                SaveSettings();
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
                // Create a subfolder based on the base part of the order number (before any dash).
                string baseNumber = orderNumber;
                int dashIndex = orderNumber.IndexOf('-');
                if (dashIndex > 0)
                {
                    baseNumber = orderNumber.Substring(0, dashIndex);
                }

                // Determine and create the final destination folder.  This folder is a child of the
                // selected target folder and is named after the base order number.  Creating the
                // directory here ensures subsequent file operations do not fail on a missing folder.
                string subFolder = System.IO.Path.Combine(TargetFolderTextBox.Text, baseNumber);
                try
                {
                    System.IO.Directory.CreateDirectory(subFolder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Kunne ikke opprette mappe '{subFolder}': {ex.Message}", "Mappefeil",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                ProcessRawData(subFolder, orderNumber, _selectedRawFile);
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

            // Step 2: Read raw data
            string rawFilePath = Path.Combine(_settings.RawFileLocation, rawFileName);
            var columnCData = new List<string>();
            var columnKData = new List<string>();

            StatusTextBlock.Text += $"Leser rådata fra {rawFileName}...\n";

            string extension = Path.GetExtension(rawFilePath).ToLower();

            try
            {
                if (extension == ".xls")
                {
                    // Bruk ExcelDataReader for gamle .xls-filer
                    ReadXlsFile(rawFilePath, columnCData, columnKData);
                }
                else
                {
                    // Bruk ClosedXML for .xlsx-filer
                    ReadXlsxFile(rawFilePath, columnCData, columnKData);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Feil ved lesing av råfil: {ex.Message}. Sørg for at filen ikke er åpen i et annet program.");
            }

            StatusTextBlock.Text += $"✓ Lest {columnCData.Count} rader fra kolonne C\n";
            StatusTextBlock.Text += $"✓ Lest {columnKData.Count} rader fra kolonne K\n\n";

            // Step 3: Write data to target file (F)
            StatusTextBlock.Text += $"Skriver data til {orderNumber} F.xlsx...\n";

            using (var targetWorkbook = new XLWorkbook(targetFileF))
            {
                var targetSheet = targetWorkbook.Worksheets.First();

                for (int i = 0; i < columnCData.Count; i++)
                {
                    targetSheet.Cell(i + 2, 5).Value = columnCData[i];
                }

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

            _onProcessComplete?.Invoke(targetFileF);
            Close();
        }

        /// <summary>
        /// Leser data fra gamle .xls-filer ved hjelp av ExcelDataReader
        /// </summary>
        private void ReadXlsFile(string filePath, List<string> columnCData, List<string> columnKData)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = false
                        }
                    });

                    int sheetCount = result.Tables.Count;
                    StatusTextBlock.Text += $"Fant {sheetCount} sheet(s) i råfilen\n";

                    foreach (System.Data.DataTable table in result.Tables)
                    {
                        StatusTextBlock.Text += $"  Prosesserer sheet: {table.TableName}\n";

                        // Les kolonne C (index 2, siden det er 0-basert)
                        for (int row = 2; row < table.Rows.Count; row++) // Start fra rad 3 (index 2)
                        {
                            if (table.Columns.Count > 2)
                            {
                                string valueC = table.Rows[row][2]?.ToString()?.Trim() ?? "";
                                if (string.IsNullOrEmpty(valueC))
                                    break;
                                columnCData.Add(valueC);
                            }
                        }

                        // Les kolonne K (index 10)
                        for (int row = 2; row < table.Rows.Count; row++) // Start fra rad 3 (index 2)
                        {
                            if (table.Columns.Count > 10)
                            {
                                string valueK = table.Rows[row][10]?.ToString()?.Trim() ?? "";
                                if (string.IsNullOrEmpty(valueK))
                                    break;
                                columnKData.Add(valueK);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Leser data fra .xlsx-filer ved hjelp av ClosedXML
        /// </summary>
        private void ReadXlsxFile(string filePath, List<string> columnCData, List<string> columnKData)
        {
            using (var rawWorkbook = new XLWorkbook(filePath))
            {
                int sheetCount = rawWorkbook.Worksheets.Count;
                StatusTextBlock.Text += $"Fant {sheetCount} sheet(s) i råfilen\n";

                foreach (var sheet in rawWorkbook.Worksheets)
                {
                    StatusTextBlock.Text += $"  Prosesserer sheet: {sheet.Name}\n";

                    // Les kolonne C fra rad 3 og nedover
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

                    // Les kolonne K fra rad 3 og nedover
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
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /// <summary>
        /// When the order number textbox changes, automatically populate the search box with the base
        /// portion of the order number (the part before any dash).  This allows the user to simply
        /// enter the full order number once and immediately filter the raw file list without
        /// retyping the base number.  The caret is moved to the end of the search box so further
        /// typing continues to append to the base.
        /// </summary>
        private void OrderNumberTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string order = OrderNumberTextBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(order))
            {
                SearchTextBox.Text = string.Empty;
                return;
            }
            string basePart = order;
            int dash = order.IndexOf('-');
            if (dash > 0)
            {
                basePart = order.Substring(0, dash);
            }
            // Only update if the search box does not already have the same basePart to avoid
            // interfering with user modifications.  Using IndexOf ensures that additional
            // characters typed by the user are preserved once the base portion has been set.
            if (!SearchTextBox.Text.Equals(basePart, StringComparison.Ordinal))
            {
                SearchTextBox.Text = basePart;
                SearchTextBox.CaretIndex = SearchTextBox.Text.Length;
            }
        }
    }

    public class RawDataSettings
    {
        public string RawFileLocation { get; set; } = "";
        public string TemplateFile1 { get; set; } = ""; // Fil som får data (F)
        public string TemplateFile2 { get; set; } = ""; // 4. Nord Ledningsliste - mal (N)
        public string TemplateFile3 { get; set; } = ""; // 3. Durapart ledningsliste - mal (D)

        // Stores the last target folder used when processing raw data.  Persisting this value
        // allows the user to avoid re‑selecting the destination folder each time.  When not set
        // or empty, the target folder text box will remain blank until the user chooses one.
        public string LastTargetFolder { get; set; } = "";
    }
}