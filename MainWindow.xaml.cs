using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace FilterV1
{
    public partial class MainWindow : Window
    {
        // Reference to the custom cross-section window. Opening it modelessly allows the user
        // to copy data from the data preview while the window is open. Only one instance is
        // maintained at a time.
        private CustomCrossSectionWindow _customCrossWindow;
        private string _filePath;
        private DataTable _dataTable;
        private Stack<DataTable> _undoStack;
        private int _rowsRemoved;
        private bool _removeEqualsApplied;
        private bool _removeLVApplied;
        private List<GroupDefinition> _customGroups;
        private List<TextFillPattern> _textFillPatterns;
        private List<StarDupesRule> _starDupesRules;
        private List<RemoveRelayPattern> _removeRelayPatterns;
        private List<ConversionRule> _conversionRules;
        private readonly string _settingsFilePath;
        private readonly string _textFillSettingsFilePath;
        private readonly string _starDupesSettingsFilePath;
        private readonly string _removeRelaySettingsFilePath;
        private readonly string _conversionRulesSettingsFilePath;
        // File path for storing the rising numbers exception list. Persisted between sessions.
        private readonly string _risingExceptionsFilePath;

        // Added for rising numbers exceptions and custom cross‑section handling
        private List<string> _risingNumberExceptions = new List<string>();
        private List<CustomCrossSectionWindow.CrossRow> _customCrossRows = new List<CustomCrossSectionWindow.CrossRow>();
        private HashSet<int> _rowsWithCustomCross = new HashSet<int>();

        public MainWindow()
        {
            InitializeComponent();

            _undoStack = new Stack<DataTable>();
            _rowsRemoved = 0;
            _removeEqualsApplied = false;
            _removeLVApplied = false;

            // Set up settings file paths
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appFolder = Path.Combine(appDataPath, "FilterV1");
            Directory.CreateDirectory(appFolder);
            _settingsFilePath = Path.Combine(appFolder, "GroupSettings.json");
            _textFillSettingsFilePath = Path.Combine(appFolder, "TextFillSettings.json");
            _starDupesSettingsFilePath = Path.Combine(appFolder, "StarDupesSettings.json");
            _removeRelaySettingsFilePath = Path.Combine(appFolder, "RemoveRelaySettings.json");
            _conversionRulesSettingsFilePath = Path.Combine(appFolder, "ConversionRulesSettings.json");

            // Set up the file path for rising number exceptions.  This list will be loaded
            // during startup and saved whenever the user modifies it via the options window.
            _risingExceptionsFilePath = Path.Combine(appFolder, "RisingExceptions.json");

            // After all file paths are set, load previously saved exceptions for the rising numbers removal.
            // This ensures that user-defined exceptions persist across sessions.  The file path
            // must be initialized before calling this method.
            LoadRisingExceptions();

            // Load all settings from persistent storage
            LoadCustomGroups();
            LoadTextFillPatterns();
            LoadStarDupesRules();
            LoadRemoveRelayPatterns();
            LoadConversionRules();
        }

        private void LoadCustomGroups()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    string json = File.ReadAllText(_settingsFilePath);
                    _customGroups = JsonSerializer.Deserialize<List<GroupDefinition>>(json) ?? GetDefaultGroups();
                }
                else
                {
                    _customGroups = GetDefaultGroups();
                    SaveCustomGroups();
                }
            }
            catch (Exception)
            {
                _customGroups = GetDefaultGroups();
            }
        }

        private void SaveCustomGroups()
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(_customGroups, options);
                File.WriteAllText(_settingsFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save group settings: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void LoadTextFillPatterns()
        {
            try
            {
                if (File.Exists(_textFillSettingsFilePath))
                {
                    string json = File.ReadAllText(_textFillSettingsFilePath);
                    _textFillPatterns = JsonSerializer.Deserialize<List<TextFillPattern>>(json) ?? new List<TextFillPattern>();
                }
                else
                {
                    _textFillPatterns = new List<TextFillPattern>();
                    SaveTextFillPatterns();
                }
            }
            catch (Exception)
            {
                _textFillPatterns = new List<TextFillPattern>();
            }
        }

        private void SaveTextFillPatterns()
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(_textFillPatterns, options);
                File.WriteAllText(_textFillSettingsFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save text fill settings: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void LoadStarDupesRules()
        {
            try
            {
                if (File.Exists(_starDupesSettingsFilePath))
                {
                    string json = File.ReadAllText(_starDupesSettingsFilePath);
                    _starDupesRules = JsonSerializer.Deserialize<List<StarDupesRule>>(json) ?? GetDefaultStarDupesRules();
                }
                else
                {
                    _starDupesRules = GetDefaultStarDupesRules();
                    SaveStarDupesRules();
                }
            }
            catch (Exception)
            {
                _starDupesRules = GetDefaultStarDupesRules();
            }
        }

        private void SaveStarDupesRules()
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(_starDupesRules, options);
                File.WriteAllText(_starDupesSettingsFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save star dupes settings: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void LoadRemoveRelayPatterns()
        {
            try
            {
                if (File.Exists(_removeRelaySettingsFilePath))
                {
                    string json = File.ReadAllText(_removeRelaySettingsFilePath);
                    _removeRelayPatterns = JsonSerializer.Deserialize<List<RemoveRelayPattern>>(json) ?? new List<RemoveRelayPattern>();
                }
                else
                {
                    _removeRelayPatterns = new List<RemoveRelayPattern>();
                    SaveRemoveRelayPatterns();
                }
            }
            catch (Exception)
            {
                _removeRelayPatterns = new List<RemoveRelayPattern>();
            }
        }

        private void SaveRemoveRelayPatterns()
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(_removeRelayPatterns, options);
                File.WriteAllText(_removeRelaySettingsFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save remove relay settings: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void LoadConversionRules()
        {
            try
            {
                if (File.Exists(_conversionRulesSettingsFilePath))
                {
                    string json = File.ReadAllText(_conversionRulesSettingsFilePath);
                    _conversionRules = JsonSerializer.Deserialize<List<ConversionRule>>(json) ?? new List<ConversionRule>();
                }
                else
                {
                    _conversionRules = new List<ConversionRule>();
                    SaveConversionRules();
                }
            }
            catch (Exception)
            {
                _conversionRules = new List<ConversionRule>();
            }
        }

        private void SaveConversionRules()
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(_conversionRules, options);
                File.WriteAllText(_conversionRulesSettingsFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save conversion rules settings: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private List<GroupDefinition> GetDefaultGroups()
        {
            return new List<GroupDefinition>
            {
                new GroupDefinition { GroupName = "X1 Group", ContainsText = "X1:", Priority = 1 },
                new GroupDefinition { GroupName = "F11 Group", ContainsText = "F11-X", Priority = 2 }
            };
        }

        private List<StarDupesRule> GetDefaultStarDupesRules()
        {
            return new List<StarDupesRule>
            {
                new StarDupesRule { DuplicateContains = "X2", AdjacentContains = "X1", Priority = 1 }
            };
        }

        private void UploadButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _filePath = openFileDialog.FileName;
                FilePathText.Text = Path.GetFileName(_filePath);
                _undoStack.Clear();
                _rowsRemoved = 0;
                _removeEqualsApplied = false;
                _removeLVApplied = false;
                LoadExcelFile();
            }
        }

        private void LoadExcelFile()
        {
            try
            {
                using (var workbook = new XLWorkbook(_filePath))
                {
                    var worksheet = workbook.Worksheets.First();
                    _dataTable = new DataTable();

                    foreach (var cell in worksheet.Row(1).CellsUsed())
                    {
                        _dataTable.Columns.Add(cell.GetString());
                    }

                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        var dataRow = _dataTable.NewRow();
                        for (int i = 0; i < _dataTable.Columns.Count; i++)
                        {
                            dataRow[i] = row.Cell(i + 1).GetString().Trim();
                        }
                        _dataTable.Rows.Add(dataRow);
                    }

                    UpdateGrid($"Fil lastet successfully. Total rader: {_dataTable.Rows.Count}");

                    // Show DataGrid and hide empty state
                    EmptyState.Visibility = Visibility.Collapsed;
                    ExcelDataGrid.Visibility = Visibility.Visible;
                    ExcelDataGrid.ItemsSource = _dataTable.DefaultView;
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Feil ved lasting av fil - {ex.Message}";
            }
        }

        private void SaveState()
        {
            if (_dataTable != null)
            {
                DataTable clone = _dataTable.Copy();
                _undoStack.Push(clone);
            }
        }

        private void MoveCellsUpward()
        {
            if (_dataTable == null)
                return;

            int initialRowCount = _dataTable.Rows.Count;

            for (int col = 0; col < _dataTable.Columns.Count; col++)
            {
                int writeRow = 0;
                for (int row = 0; row < _dataTable.Rows.Count; row++)
                {
                    string cellValue = _dataTable.Rows[row][col]?.ToString();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        if (writeRow != row)
                        {
                            _dataTable.Rows[writeRow][col] = cellValue;
                            _dataTable.Rows[row][col] = string.Empty;
                        }
                        writeRow++;
                    }
                }
            }

            for (int i = _dataTable.Rows.Count - 1; i >= 0; i--)
            {
                bool isEmpty = true;
                for (int j = 0; j < _dataTable.Columns.Count; j++)
                {
                    if (!string.IsNullOrEmpty(_dataTable.Rows[i][j]?.ToString()))
                    {
                        isEmpty = false;
                        break;
                    }
                }
                if (isEmpty)
                {
                    _dataTable.Rows.RemoveAt(i);
                }
            }

            _rowsRemoved += initialRowCount - _dataTable.Rows.Count;
        }

        private void UpdateGrid(string statusMessage)
        {
            ExcelDataGrid.ItemsSource = null;
            ExcelDataGrid.ItemsSource = _dataTable?.DefaultView;
            StatusText.Text = $"{statusMessage} (Fjernet rader: {_rowsRemoved}, Total rader: {_dataTable?.Rows.Count ?? 0})";
            RowCountText.Text = $"{_dataTable?.Rows.Count ?? 0} rader";

            if (_dataTable != null && _dataTable.Rows.Count > 0)
            {
                EmptyState.Visibility = Visibility.Collapsed;
                ExcelDataGrid.Visibility = Visibility.Visible;
            }
            else
            {
                EmptyState.Visibility = Visibility.Visible;
                ExcelDataGrid.Visibility = Visibility.Collapsed;
            }
        }

        private void RemoveAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            SaveState();

            // Execute all removal functions in sequence (original buttons 1, 2, 3, 4)

            // 1. Remove = from cells
            foreach (DataRow row in _dataTable.Rows)
            {
                for (int i = 0; i < _dataTable.Columns.Count; i++)
                {
                    string cellValue = row[i]?.ToString();
                    if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains("="))
                    {
                        row[i] = cellValue.Replace("=", "");
                    }
                }
            }

            // 2. Remove +LV from cells
            foreach (DataRow row in _dataTable.Rows)
            {
                for (int i = 0; i < _dataTable.Columns.Count; i++)
                {
                    string cellValue = row[i]?.ToString();
                    if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains("+LV"))
                    {
                        row[i] = cellValue.Replace("+LV", "");
                    }
                }
            }

            // 3. Remove MX and adjacent cells
            foreach (DataRow row in _dataTable.Rows)
            {
                for (int j = 0; j < _dataTable.Columns.Count; j++)
                {
                    string cellValue = row[j]?.ToString();
                    if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains("MX"))
                    {
                        row[j] = string.Empty;
                        if (j < _dataTable.Columns.Count - 1)
                        {
                            row[j + 1] = string.Empty;
                        }
                        if (j > 0)
                        {
                            row[j - 1] = string.Empty;
                        }
                    }
                }
            }

            // 4. Remove XS and adjacent cells
            foreach (DataRow row in _dataTable.Rows)
            {
                for (int j = 0; j < _dataTable.Columns.Count; j++)
                {
                    string cellValue = row[j]?.ToString();
                    if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains("XS"))
                    {
                        row[j] = string.Empty;
                        if (j > 0)
                        {
                            row[j - 1] = string.Empty;
                        }
                        if (j < _dataTable.Columns.Count - 1)
                        {
                            row[j + 1] = string.Empty;
                        }
                    }
                }
            }

            // Move cells upward after all removals
            MoveCellsUpward();
            _removeEqualsApplied = true;
            _removeLVApplied = true;
            UpdateGrid("Fjernet =, +LV, MX, og XS fra celler");
        }

        private void RemoveRisingNumbersButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }
            // Open the rising numbers options window to allow the user to specify exceptions.
            // Any exception strings returned will be used to skip rows where column 5 or 6 contains the exception.
            var optsWin = new RisingNumbersOptionsWindow(_risingNumberExceptions, ex =>
            {
                // Update the exception list with the selection from the dialog. Use an empty
                // list instead of null to avoid null reference checks later.
                _risingNumberExceptions = ex ?? new List<string>();
            });
            optsWin.Owner = this;
            optsWin.ShowDialog();
            // Persist the current exception list to disk so it is remembered between sessions.
            SaveRisingExceptions();

            SaveState();
            var clearedCells = new List<string>();
            for (int ri = 0; ri < _dataTable.Rows.Count; ri++)
            {
                DataRow row = _dataTable.Rows[ri];
                string col5Value = row[4]?.ToString();
                string col6Value = row[5]?.ToString();

                // If either column contains an exception string (case‑insensitive), skip this row.
                bool skip = false;
                if (_risingNumberExceptions != null && _risingNumberExceptions.Any())
                {
                    foreach (string ex in _risingNumberExceptions)
                    {
                        if (string.IsNullOrWhiteSpace(ex)) continue;
                        if ((!string.IsNullOrEmpty(col5Value) && col5Value.IndexOf(ex, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (!string.IsNullOrEmpty(col6Value) && col6Value.IndexOf(ex, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            skip = true;
                            break;
                        }
                    }
                }
                if (skip) continue;

                if (!string.IsNullOrEmpty(col5Value) && !string.IsNullOrEmpty(col6Value))
                {
                    var match5 = Regex.Match(col5Value, @"^(.*):(\d+)$");
                    var match6 = Regex.Match(col6Value, @"^(.*):(\d+)$");
                    if (match5.Success && match6.Success)
                    {
                        string prefix5 = match5.Groups[1].Value;
                        string prefix6 = match6.Groups[1].Value;
                        if (prefix5 == prefix6)
                        {
                            if (int.TryParse(match5.Groups[2].Value, out int number5) && int.TryParse(match6.Groups[2].Value, out int number6))
                            {
                                // Check both directions for rising numbers
                                if (number6 == number5 + 1)
                                {
                                    row[4] = string.Empty;
                                    row[5] = string.Empty;
                                    clearedCells.Add($"Rad {ri + 1}: {col5Value} → {col6Value}");
                                }
                                else if (number5 == number6 + 1)
                                {
                                    row[4] = string.Empty;
                                    row[5] = string.Empty;
                                    clearedCells.Add($"Rad {ri + 1}: {col6Value} → {col5Value}");
                                }
                            }
                        }
                    }
                }
            }

            MoveCellsUpward();
            UpdateGrid($"Stigende tall fjernet (begge retninger): {string.Join(", ", clearedCells.Take(5))}{(clearedCells.Count > 5 ? "..." : "")}");
        }

        private void RemoveDefinedCellsButton_Click(object sender, RoutedEventArgs e)
        {
            var window = new RemoveCellsWindow(pairs =>
            {
                if (_dataTable != null)
                {
                    SaveState();
                    var matchedRows = new List<int>();
                    for (int i = _dataTable.Rows.Count - 1; i >= 0; i--)
                    {
                        string col5Value = _dataTable.Rows[i][4]?.ToString()?.Trim();
                        string col6Value = _dataTable.Rows[i][5]?.ToString()?.Trim();
                        if (string.IsNullOrEmpty(col5Value) || string.IsNullOrEmpty(col6Value))
                            continue;

                        foreach (var pair in pairs)
                        {
                            if (string.IsNullOrEmpty(pair.FirstCell) || string.IsNullOrEmpty(pair.SecondCell))
                                continue;

                            // FORBEDRET LOGIKK: Sjekk begge kombinasjoner i samme rad
                            bool pattern1Match = (col5Value.Contains(pair.FirstCell, StringComparison.OrdinalIgnoreCase) &&
                                                col6Value.Contains(pair.SecondCell, StringComparison.OrdinalIgnoreCase)) ||
                                               (col5Value.Contains(pair.SecondCell, StringComparison.OrdinalIgnoreCase) &&
                                                col6Value.Contains(pair.FirstCell, StringComparison.OrdinalIgnoreCase));

                            if (pattern1Match)
                            {
                                matchedRows.Add(i + 1);
                                _dataTable.Rows.RemoveAt(i);
                                break;
                            }
                        }
                    }
                    MoveCellsUpward();
                    UpdateGrid($"Fjernet rader med matchende mønstre i samme rad (kol 5-6): {string.Join(", ", matchedRows.OrderByDescending(x => x))}");
                }
                else
                {
                    UpdateGrid("Ingen fil lastet, mønstre definert men ikke anvendt");
                }
            });
            window.Show();
        }

        private void RemoveDupesButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            SaveState();
            foreach (DataRow row in _dataTable.Rows)
            {
                for (int j = 0; j < _dataTable.Columns.Count - 1; j++)
                {
                    string leftCell = row[j]?.ToString();
                    string rightCell = row[j + 1]?.ToString();
                    if (!string.IsNullOrEmpty(leftCell) && leftCell == rightCell)
                    {
                        row[j] = string.Empty;
                        row[j + 1] = string.Empty;
                    }
                }
            }
            MoveCellsUpward();
            UpdateGrid("Duplikat tilstøtende celler fjernet");
        }

        private void SortButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            // Open the custom group window with current groups
            var window = new CustomGroupWindow(_customGroups, groupDefinitions =>
            {
                _customGroups = groupDefinitions;
                SaveCustomGroups(); // Save to persistent storage
                ApplyCustomSort();
            });
            window.Owner = this;
            // Open modelessly so the user can continue interacting with the main window
            window.Show();
        }

        /// <summary>
        /// Opens the custom cross‑section window for defining per‑row tverrsnitt assignments.
        /// After the dialog is closed, applies the definitions to the current data table.
        /// </summary>
        /// <summary>
        /// Opens the custom cross section window. Before opening, attempts to parse the current clipboard contents into
        /// a list of cross rows if the clipboard contains tabular text (e.g. copied from Excel or the app's DataGrid).
        /// The parsed rows will seed the window so the user can immediately apply tverrsnitt options without manually pasting.
        /// </summary>
        private void CustomCrossSectionButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }
            // If the custom cross‑section window is already open, just bring it to the front.
            if (_customCrossWindow != null && _customCrossWindow.IsLoaded)
            {
                _customCrossWindow.Activate();
                return;
            }

            // Create and show a modeless custom cross‑section window.  This allows the user to
            // interact with the main window (for example, to copy values from the data preview)
            // while defining custom mappings.  When the window closes, the reference is cleared
            // and the custom mappings are applied.
            _customCrossWindow = new CustomCrossSectionWindow(_customCrossRows, rows =>
            {
                _customCrossRows = rows ?? new System.Collections.Generic.List<CustomCrossSectionWindow.CrossRow>();
                ApplyCustomCrossSections();
            });
            _customCrossWindow.Owner = this;
            _customCrossWindow.Closed += (_, __) => _customCrossWindow = null;
            _customCrossWindow.Show();
        }

        /// <summary>
        /// Applies the custom cross‑section definitions stored in _customCrossRows to the
        /// current DataTable. For each row in the table, if the values in columns 5 and 6
        /// match a defined pattern, the tverrsnitt option is applied to columns 2–4.
        /// Rows that receive a custom tverrsnitt are recorded in _rowsWithCustomCross
        /// so that standard tverrsnitt rules will not overwrite them.
        /// </summary>
        private void ApplyCustomCrossSections()
        {
            if (_dataTable == null) return;

            SaveState();
            _rowsWithCustomCross.Clear();

            for (int i = 0; i < _dataTable.Rows.Count; i++)
            {
                DataRow row = _dataTable.Rows[i];
                string col5 = row[4]?.ToString() ?? string.Empty;
                string col6 = row[5]?.ToString() ?? string.Empty;

                foreach (var def in _customCrossRows)
                {
                    if (def == null) continue;
                    bool has5 = !string.IsNullOrWhiteSpace(def.Col5Text);
                    bool has6 = !string.IsNullOrWhiteSpace(def.Col6Text);

                    // Direct order: match col5→def.Col5Text and col6→def.Col6Text if specified
                    bool matchDirect = (!has5 || (col5.IndexOf(def.Col5Text, StringComparison.OrdinalIgnoreCase) >= 0)) &&
                                       (!has6 || (col6.IndexOf(def.Col6Text, StringComparison.OrdinalIgnoreCase) >= 0));

                    // Swapped order: match col6→def.Col5Text and col5→def.Col6Text if specified
                    bool matchSwapped = (!has5 || (col6.IndexOf(def.Col5Text, StringComparison.OrdinalIgnoreCase) >= 0)) &&
                                        (!has6 || (col5.IndexOf(def.Col6Text, StringComparison.OrdinalIgnoreCase) >= 0));

                    if (matchDirect || matchSwapped)
                    {
                        var (c2, c3, c4) = MapCrossOption(def.SelectedOption);
                        row[1] = c2;
                        row[2] = c3;
                        row[3] = c4;
                        _rowsWithCustomCross.Add(i);
                        break; // Only first matching definition applies
                    }
                }
            }

            UpdateGrid("Anvendt egendefinert tverrsnitt");
        }

        /// <summary>
        /// Maps a cross option index to the three corresponding tverrsnitt values for columns 2, 3, and 4.
        /// </summary>
        /// <param name="option">Option index (1–4)</param>
        /// <returns>Tuple of (col2, col3, col4) strings</returns>
        private (string c2, string c3, string c4) MapCrossOption(int option)
        {
            switch (option)
            {
                case 2: return ("UNIBK1.5", "HYLSE 1.5", "HYLSE 1.5");
                case 3: return ("UNIBK2.5", "HYLSE 2.5", "HYLSE 2.5");
                case 4: return ("UNIBK4.0", "HYLSE 4.0", "HYLSE 4.0");
                default: return ("UNIBK1.0", "HYLSE 1.0", "HYLSE 1.0");
            }
        }

        /// <summary>
        /// Saves the current list of rising numbers exceptions to a JSON file. This ensures
        /// that the user's exception preferences persist across sessions. Errors are
        /// swallowed silently because failure to save should not prevent the user from
        /// continuing to use the application.
        /// </summary>
        private void SaveRisingExceptions()
        {
            try
            {
                // Ensure the directory exists. GetDirectoryName returns null if the path
                // contains no directory information, but our path always includes a folder.
                var dir = Path.GetDirectoryName(_risingExceptionsFilePath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(_risingNumberExceptions ?? new List<string>(), options);
                File.WriteAllText(_risingExceptionsFilePath, json);
            }
            catch
            {
                // Ignore any exceptions when saving. The app can still function without persistence.
            }
        }

        /// <summary>
        /// Loads the rising numbers exceptions from the JSON file if it exists. If the file
        /// cannot be read or does not exist, the list is reset to an empty list. This
        /// method should be called during application startup.
        /// </summary>
        private void LoadRisingExceptions()
        {
            try
            {
                if (File.Exists(_risingExceptionsFilePath))
                {
                    string json = File.ReadAllText(_risingExceptionsFilePath);
                    var list = JsonSerializer.Deserialize<List<string>>(json);
                    _risingNumberExceptions = list ?? new List<string>();
                }
                else
                {
                    _risingNumberExceptions = new List<string>();
                }
            }
            catch
            {
                // In case of any deserialization error, fall back to an empty list.
                _risingNumberExceptions = new List<string>();
            }
        }

        private void ConvertToDurapartButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            // Open the convert to durapart window with current rules
            var window = new ConvertToDurapartWindow(_conversionRules, rules =>
            {
                _conversionRules = rules;
                SaveConversionRules(); // Save to persistent storage
                ApplyConversionRules();
            });
            window.Owner = this;
            window.ShowDialog();
        }

        private void ReorganizeCellsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            SaveState();
            ApplyReorganizeCells();
        }

        private void CustomTextFillButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            // Open the custom text fill window with current patterns
            var window = new CustomTextFillWindow(_textFillPatterns, patterns =>
            {
                _textFillPatterns = patterns;
                SaveTextFillPatterns(); // Save to persistent storage
                ApplyTextFillPatterns();
            });
            window.Owner = this;
            window.ShowDialog();
        }

        private void StarDupesButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            // Open the star dupes window with current rules
            var window = new StarDupesWindow(_starDupesRules, rules =>
            {
                _starDupesRules = rules;
                SaveStarDupesRules(); // Save to persistent storage
                ApplyStarDupes();
            });
            window.Owner = this;
            // Show modelessly so the user can copy/paste from the main window while defining rules
            window.Show();
        }

        private void RemoveRelayButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            // Open the remove relay window with current patterns
            var window = new RemoveRelayWindow(_removeRelayPatterns, (allPatterns, enabledPatterns) =>
            {
                // Save all patterns for persistence, but only process enabled ones
                _removeRelayPatterns = allPatterns; // Save all patterns for next time
                SaveRemoveRelayPatterns(); // Save to persistent storage

                // Process only enabled patterns
                ApplyRemoveRelay(enabledPatterns);
            });
            window.Owner = this;
            window.ShowDialog();
        }

        private void ApplyTextFillPatterns()
        {
            if (_dataTable == null) return;

            SaveState();
            var modifiedRows = new List<int>();

            for (int i = 0; i < _dataTable.Rows.Count; i++)
            {
                // Skip rows that have been assigned a custom cross‑section
                if (_rowsWithCustomCross.Contains(i))
                    continue;

                DataRow row = _dataTable.Rows[i];
                string col5Value = row[4]?.ToString()?.Trim() ?? string.Empty;
                string col6Value = row[5]?.ToString()?.Trim() ?? string.Empty;

                foreach (var pattern in _textFillPatterns.OrderBy(p => p.Priority))
                {
                    if (!string.IsNullOrEmpty(pattern.ContainsText) &&
                        ((col5Value.IndexOf(pattern.ContainsText, StringComparison.OrdinalIgnoreCase) >= 0) ||
                         (col6Value.IndexOf(pattern.ContainsText, StringComparison.OrdinalIgnoreCase) >= 0)))
                    {
                        int rowIndex = i + 1;
                        var (col2, col3, col4) = GetOptionValues(pattern.SelectedOption);
                        row[1] = col2;
                        row[2] = col3;
                        row[3] = col4;
                        modifiedRows.Add(rowIndex);
                        break; // apply only first matching pattern
                    }
                }
            }

            UpdateGrid($"Anvendt tekst utfylling til kolonne 2-4 i rader: {string.Join(", ", modifiedRows.OrderBy(x => x))}");
        }

        private (string col2, string col3, string col4) GetOptionValues(int option)
        {
            switch (option)
            {
                case 1: return ("UNIBK1.0", "HYLSE 1.0", "HYLSE 1.0");
                case 2: return ("UNIBK1.5", "HYLSE 1.5", "HYLSE 1.5");
                case 3: return ("UNIBK2.5", "HYLSE 2.5", "HYLSE 2.5");
                case 4: return ("UNIBK4.0", "HYLSE 4.0", "HYLSE 4.0");
                default: return ("UNIBK1.0", "HYLSE 1.0", "HYLSE 1.0");
            }
        }

        private void ApplyCustomSort()
        {
            if (_dataTable == null) return;

            SaveState();

            // IMPORTANT: Ensure higher priority text is always on the left (column 5)
            if (_dataTable.Columns.Count >= 6)
            {
                for (int i = 0; i < _dataTable.Rows.Count; i++)
                {
                    string col5Value = _dataTable.Rows[i][4]?.ToString() ?? "";
                    string col6Value = _dataTable.Rows[i][5]?.ToString() ?? "";

                    // Find the priority of each column's text (lower number = higher priority)
                    int col5Priority = GetTextPriority(col5Value);
                    int col6Priority = GetTextPriority(col6Value);

                    // If column 6 has higher priority text (lower number) than column 5, swap them
                    // OR if column 6 has priority text and column 5 doesn't have any priority text
                    bool shouldSwap = false;

                    if (col6Priority > 0 && col5Priority == 0)
                    {
                        // Column 6 has priority text, column 5 doesn't
                        shouldSwap = true;
                    }
                    else if (col6Priority > 0 && col5Priority > 0 && col6Priority < col5Priority)
                    {
                        // Both have priority text, but column 6 has higher priority (lower number)
                        shouldSwap = true;
                    }

                    if (shouldSwap)
                    {
                        _dataTable.Rows[i][4] = col6Value;  // Move higher priority text to column 5
                        _dataTable.Rows[i][5] = col5Value;  // Move lower priority text to column 6
                    }
                }
            }

            // Group rows based on custom definitions with priority-based assignment
            var groupedRows = new List<(GroupDefinition Group, List<DataRow> Rows)>();
            var ungroupedRows = new List<DataRow>();

            // Sort groups by priority (ascending - lower number = higher priority)
            var sortedGroups = _customGroups.OrderBy(g => g.Priority).ToList();

            // Track which rows have been assigned to prevent double-assignment
            var assignedRows = new HashSet<DataRow>();

            // Process each group definition in priority order (highest priority first)
            foreach (var group in sortedGroups)
            {
                var matchingRows = new List<DataRow>();

                for (int i = 0; i < _dataTable.Rows.Count; i++)
                {
                    var row = _dataTable.Rows[i];

                    // Skip if this row is already assigned to a higher priority group
                    if (assignedRows.Contains(row)) continue;

                    string col5Value = row[4]?.ToString() ?? "";
                    string col6Value = row[5]?.ToString() ?? "";

                    // Check if this row matches the current group (check both columns 5 and 6)
                    if (col5Value.Contains(group.ContainsText) || col6Value.Contains(group.ContainsText))
                    {
                        matchingRows.Add(row);
                        assignedRows.Add(row); // Mark as assigned
                    }
                }

                if (matchingRows.Count > 0)
                {
                    // Sort within group
                    matchingRows.Sort((a, b) =>
                    {
                        string aValue = a[4]?.ToString() ?? "";
                        string bValue = b[4]?.ToString() ?? "";
                        return string.Compare(aValue, bValue, StringComparison.Ordinal);
                    });

                    groupedRows.Add((group, matchingRows));
                }
            }

            // Collect ungrouped rows (rows that don't match any group definition)
            foreach (DataRow row in _dataTable.Rows)
            {
                if (!assignedRows.Contains(row))
                {
                    ungroupedRows.Add(row);
                }
            }

            // Create new sorted table
            DataTable sortedTable = _dataTable.Clone();

            // Add grouped rows with sequential group numbers (1, 2, 3, 4...)
            int groupIndex = 1;
            foreach (var (group, rows) in groupedRows)
            {
                foreach (var row in rows)
                {
                    var newRow = sortedTable.NewRow();
                    newRow.ItemArray = row.ItemArray;

                    // Set group number for the entire group
                    if (sortedTable.Columns.Count > 6)
                    {
                        newRow[6] = groupIndex.ToString();
                    }
                    else if (sortedTable.Columns.Count == 6)
                    {
                        sortedTable.Columns.Add("Group", typeof(string));
                        newRow[6] = groupIndex.ToString();
                    }

                    sortedTable.Rows.Add(newRow);
                }
                groupIndex++; // Move to next group number
            }

            // Add ungrouped rows at the end
            foreach (var row in ungroupedRows)
            {
                var newRow = sortedTable.NewRow();
                newRow.ItemArray = row.ItemArray;

                if (sortedTable.Columns.Count > 6)
                {
                    newRow[6] = ""; // No group
                }

                sortedTable.Rows.Add(newRow);
            }

            _dataTable = sortedTable;
            MoveCellsUpward();

            var activeGroups = groupedRows.Select(g => g.Group.ContainsText).ToList();
            string groupSummary = string.Join(", ", activeGroups);
            UpdateGrid($"Data sortert med prioritet tekst flyttet til venstre. Grupper: {groupSummary}");
        }

        private void ApplyStarDupes()
        {
            if (_dataTable == null) return;

            SaveState();
            var modifiedCells = new List<string>();

            // Find all values in columns 5 and 6
            var cellValues = new Dictionary<string, List<(int rowIndex, int colIndex, string adjacentValue)>>();

            for (int i = 0; i < _dataTable.Rows.Count; i++)
            {
                // Check column 5
                string col5Value = _dataTable.Rows[i][4]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(col5Value) && !col5Value.EndsWith("*"))
                {
                    string col6Adjacent = _dataTable.Rows[i][5]?.ToString()?.Trim() ?? "";
                    if (!cellValues.ContainsKey(col5Value))
                        cellValues[col5Value] = new List<(int, int, string)>();
                    cellValues[col5Value].Add((i, 4, col6Adjacent)); // column 4 = column 5 (0-indexed)
                }

                // Check column 6
                string col6Value = _dataTable.Rows[i][5]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(col6Value) && !col6Value.EndsWith("*"))
                {
                    string col5Adjacent = _dataTable.Rows[i][4]?.ToString()?.Trim() ?? "";
                    if (!cellValues.ContainsKey(col6Value))
                        cellValues[col6Value] = new List<(int, int, string)>();
                    cellValues[col6Value].Add((i, 5, col5Adjacent)); // column 5 = column 6 (0-indexed)
                }
            }

            // Process duplicates using the new rule-based approach
            foreach (var kvp in cellValues.Where(x => x.Value.Count > 1))
            {
                string duplicateValue = kvp.Key;
                var locations = kvp.Value;

                // Find the location that matches the highest priority rule
                (int rowIndex, int colIndex, string adjacentValue)? bestLocation = null;
                int bestPriority = int.MaxValue;

                foreach (var location in locations)
                {
                    // Check if this location matches any rule
                    foreach (var rule in _starDupesRules.OrderBy(r => r.Priority))
                    {
                        // Check if duplicate contains the rule's duplicate pattern
                        // AND the adjacent cell contains the rule's adjacent pattern
                        if (duplicateValue.Contains(rule.DuplicateContains) &&
                            location.adjacentValue.Contains(rule.AdjacentContains))
                        {
                            if (rule.Priority < bestPriority)
                            {
                                bestPriority = rule.Priority;
                                bestLocation = location;
                            }
                            break; // Found a matching rule for this location
                        }
                    }
                }

                // Mark the best location with "*" if one was found
                if (bestLocation.HasValue)
                {
                    _dataTable.Rows[bestLocation.Value.rowIndex][bestLocation.Value.colIndex] = duplicateValue + "*";
                    modifiedCells.Add($"Rad {bestLocation.Value.rowIndex + 1}, Kolonne {bestLocation.Value.colIndex + 1}");
                }
            }

            UpdateGrid($"Anvendt stjerne markering til {modifiedCells.Count} duplikat celler: {string.Join(", ", modifiedCells.Take(10))}{(modifiedCells.Count > 10 ? "..." : "")}");
        }

        private void ApplyRemoveRelay(List<RemoveRelayPattern> patternsToApply = null)
        {
            if (_dataTable == null) return;

            SaveState();
            var modifiedCells = new List<string>();

            // Use provided patterns or default to all stored patterns
            var patterns = patternsToApply ?? _removeRelayPatterns;

            // Go through all cells in the data table
            foreach (DataRow row in _dataTable.Rows)
            {
                for (int j = 0; j < _dataTable.Columns.Count; j++)
                {
                    string cellValue = row[j]?.ToString();
                    if (string.IsNullOrEmpty(cellValue)) continue;

                    // Check if this cell contains any of the remove relay patterns
                    foreach (var pattern in patterns)
                    {
                        if (cellValue.Contains(pattern.ContainsText))
                        {
                            // Remove this cell
                            modifiedCells.Add($"Rad {_dataTable.Rows.IndexOf(row) + 1}, Kol {j + 1}");
                            row[j] = string.Empty;

                            // Remove adjacent cell to the LEFT if it exists
                            if (j > 0)
                            {
                                row[j - 1] = string.Empty;
                                modifiedCells.Add($"Rad {_dataTable.Rows.IndexOf(row) + 1}, Kol {j}");
                            }

                            // Remove adjacent cell to the RIGHT if it exists
                            if (j < _dataTable.Columns.Count - 1)
                            {
                                row[j + 1] = string.Empty;
                                modifiedCells.Add($"Rad {_dataTable.Rows.IndexOf(row) + 1}, Kol {j + 2}");
                            }

                            break; // Only apply first matching pattern for this cell
                        }
                    }
                }
            }

            MoveCellsUpward();
            UpdateGrid($"Fjernet irregulære celler og tilstøtende celler (begge sider): {string.Join(", ", modifiedCells.Take(10))}{(modifiedCells.Count > 10 ? "..." : "")}");
        }

        private void ApplyReorganizeCells()
        {
            if (_dataTable == null) return;

            int totalRows = _dataTable.Rows.Count;
            int reorganizedRows = 0;

            // Check if we have enough columns (at least 6 columns needed for reorganization)
            if (_dataTable.Columns.Count < 6)
            {
                UpdateGrid("Trenger minst 6 kolonner for reorganisering");
                return;
            }

            foreach (DataRow row in _dataTable.Rows)
            {
                // Store original values
                string col2Original = row[1]?.ToString() ?? ""; // Column 2 (index 1)
                string col3Original = row[2]?.ToString() ?? ""; // Column 3 (index 2)
                string col4Original = row[3]?.ToString() ?? ""; // Column 4 (index 3)
                string col5Original = row[4]?.ToString() ?? ""; // Column 5 (index 4)
                string col6Original = row[5]?.ToString() ?? ""; // Column 6 (index 5)

                // Apply reorganization:
                // 5 → 2, 6 → 3, 2 → 4, 3 → 5, 4 → 6
                row[1] = col5Original; // Column 5 → Column 2
                row[2] = col6Original; // Column 6 → Column 3
                row[3] = col2Original; // Column 2 → Column 4
                row[4] = col3Original; // Column 3 → Column 5
                row[5] = col4Original; // Column 4 → Column 6

                reorganizedRows++;
            }

            UpdateGrid($"Reorganisert kolonner (5→2, 6→3, 2→4, 3→5, 4→6) for {reorganizedRows} rader");
        }

        private void ApplyConversionRules()
        {
            if (_dataTable == null) return;

            SaveState();
            var modifiedCells = new List<string>();

            // Go through all cells in the data table
            foreach (DataRow row in _dataTable.Rows)
            {
                for (int j = 0; j < _dataTable.Columns.Count; j++)
                {
                    string cellValue = row[j]?.ToString();
                    if (string.IsNullOrEmpty(cellValue)) continue;

                    // Check if this cell exactly matches any conversion rule
                    foreach (var rule in _conversionRules)
                    {
                        if (cellValue.Equals(rule.FromText, StringComparison.Ordinal))
                        {
                            // Replace with exact match
                            row[j] = rule.ToText;
                            modifiedCells.Add($"Rad {_dataTable.Rows.IndexOf(row) + 1}, Kol {j + 1}: '{rule.FromText}' → '{rule.ToText}'");
                            break; // Only apply first matching rule for this cell
                        }
                    }
                }
            }

            UpdateGrid($"Konvertert {modifiedCells.Count} celler til Durapart format: {string.Join(", ", modifiedCells.Take(5))}{(modifiedCells.Count > 5 ? "..." : "")}");
        }

        // Helper method to get the priority of text (returns 0 if no priority text found)
        private int GetTextPriority(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0;

            foreach (var group in _customGroups)
            {
                if (text.Contains(group.ContainsText))
                {
                    return group.Priority;
                }
            }
            return 0; // No priority text found
        }

        private void UndoButton_Click(object sender, RoutedEventArgs e)
        {
            if (_undoStack.Count == 0)
            {
                StatusText.Text = "Ingen handlinger å angre";
                return;
            }

            _dataTable = _undoStack.Pop();
            _rowsRemoved = Math.Max(0, _rowsRemoved - (_dataTable?.Rows.Count ?? 0));
            _removeEqualsApplied = _undoStack.Any() && _undoStack.Peek().Rows.Cast<DataRow>().Any(r => r.ItemArray.Any(v => v?.ToString().Contains("=") == false));
            _removeLVApplied = _undoStack.Any() && _undoStack.Peek().Rows.Cast<DataRow>().Any(r => r.ItemArray.Any(v => v?.ToString().Contains("+LV") == false));
            UpdateGrid("Siste handling angret");
        }

        /// <summary>
        /// Intercepts the Delete key in the data preview grid. When the user presses
        /// Delete to clear cell values, this handler pushes the current state onto
        /// the undo stack before any modifications occur. This allows the Undo
        /// function to restore cleared cell values. Without this handler, manual
        /// edits would not be captured by the undo stack.
        /// </summary>
        private void ExcelDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Delete)
            {
                SaveState();
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null || string.IsNullOrEmpty(_filePath))
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("FilteredData");

                    for (int i = 0; i < _dataTable.Columns.Count; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = _dataTable.Columns[i].ColumnName;
                    }

                    for (int i = 0; i < _dataTable.Rows.Count; i++)
                    {
                        var row = worksheet.Row(i + 2);
                        string colorIndexStr = _dataTable.Rows[i][6]?.ToString() ?? "";

                        if (int.TryParse(colorIndexStr, out int colorIndex))
                        {
                            // Alternate colors: odd groups = green, even groups = pink
                            XLColor rowColor = (colorIndex % 2 == 1) ? XLColor.LightGreen : XLColor.LightPink;
                            for (int j = 1; j <= _dataTable.Columns.Count; j++)
                            {
                                row.Cell(j).Style.Fill.BackgroundColor = rowColor;
                            }
                        }

                        for (int j = 0; j < _dataTable.Columns.Count; j++)
                        {
                            worksheet.Cell(i + 2, j + 1).Value = _dataTable.Rows[i][j]?.ToString();
                        }
                    }

                    workbook.SaveAs(_filePath);
                    StatusText.Text = "Fil overskrevet med farger";
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Feil ved lagring av fil - {ex.Message}";
            }
        }
    }

    public class GroupConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is string colorIndexStr && int.TryParse(colorIndexStr, out int colorIndex))
            {
                // Alternate colors: odd groups = green, even groups = pink
                return (colorIndex % 2 == 1)
                    ? new SolidColorBrush(System.Windows.Media.Color.FromArgb(51, 76, 175, 80))
                    : new SolidColorBrush(System.Windows.Media.Color.FromArgb(51, 233, 30, 99));
            }
            return Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}