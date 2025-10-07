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
        private readonly string _risingExceptionsFilePath;

        private List<string> _risingNumberExceptions = new List<string>();
        private List<CustomCrossSectionWindow.CrossRow> _customCrossRows = new List<CustomCrossSectionWindow.CrossRow>();
        private HashSet<int> _rowsWithCustomCross = new HashSet<int>();

        private bool _hasUnsavedChanges = false;

        public MainWindow()
        {
            InitializeComponent();

            this.Closing += MainWindow_Closing;

            _undoStack = new Stack<DataTable>();
            _rowsRemoved = 0;
            _removeEqualsApplied = false;
            _removeLVApplied = false;

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appFolder = Path.Combine(appDataPath, "FilterV1");
            Directory.CreateDirectory(appFolder);
            _settingsFilePath = Path.Combine(appFolder, "GroupSettings.json");
            _textFillSettingsFilePath = Path.Combine(appFolder, "TextFillSettings.json");
            _starDupesSettingsFilePath = Path.Combine(appFolder, "StarDupesSettings.json");
            _removeRelaySettingsFilePath = Path.Combine(appFolder, "RemoveRelaySettings.json");
            _conversionRulesSettingsFilePath = Path.Combine(appFolder, "ConversionRulesSettings.json");
            _risingExceptionsFilePath = Path.Combine(appFolder, "RisingExceptions.json");

            LoadRisingExceptions();
            LoadCustomGroups();
            LoadTextFillPatterns();
            LoadStarDupesRules();
            LoadRemoveRelayPatterns();
            LoadConversionRules();
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_hasUnsavedChanges && _dataTable != null)
            {
                var result = MessageBox.Show(
                    "Du har ulagrede endringer. Vil du lagre før du lukker?",
                    "Ulagrede endringer",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    SaveButton_Click(this, new RoutedEventArgs());
                }
                else if (result == MessageBoxResult.Cancel)
                {
                    e.Cancel = true;
                }
            }
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
                _hasUnsavedChanges = false;
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
                    _hasUnsavedChanges = false;

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
                _hasUnsavedChanges = true;
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
            UpdateRowCount();

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

        private void DataSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_dataTable == null || _dataTable.Rows.Count == 0)
            {
                SearchResultsText.Text = "";
                return;
            }

            string searchTerm = DataSearchTextBox.Text?.Trim() ?? "";

            if (string.IsNullOrEmpty(searchTerm))
            {
                var defaultView = _dataTable.DefaultView;
                defaultView.RowFilter = "";
                ExcelDataGrid.ItemsSource = null;
                ExcelDataGrid.ItemsSource = defaultView;
                SearchResultsText.Text = "";
                UpdateRowCount();
                return;
            }

            var filteredView = _dataTable.DefaultView;
            var filterParts = new List<string>();

            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                string columnName = _dataTable.Columns[i].ColumnName;
                string escapedSearch = searchTerm.Replace("'", "''");
                filterParts.Add($"Convert([{columnName}], 'System.String') LIKE '%{escapedSearch}%'");
            }

            try
            {
                filteredView.RowFilter = string.Join(" OR ", filterParts);
                int matchCount = filteredView.Count;

                SearchResultsText.Text = matchCount > 0
                    ? $"{matchCount} treff"
                    : "Ingen treff";

                ExcelDataGrid.ItemsSource = null;
                ExcelDataGrid.ItemsSource = filteredView;
                UpdateRowCount();
            }
            catch (Exception ex)
            {
                filteredView.RowFilter = "";
                ExcelDataGrid.ItemsSource = null;
                ExcelDataGrid.ItemsSource = filteredView;
                SearchResultsText.Text = "Søkefeil";
                UpdateRowCount();
            }
        }

        private void ClearSearchButton_Click(object sender, RoutedEventArgs e)
        {
            DataSearchTextBox.Text = "";

            if (_dataTable != null)
            {
                var defaultView = _dataTable.DefaultView;
                defaultView.RowFilter = "";
                ExcelDataGrid.ItemsSource = null;
                ExcelDataGrid.ItemsSource = defaultView;
                SearchResultsText.Text = "";
                UpdateRowCount();
            }

            DataSearchTextBox.Focus();
        }

        private void UpdateRowCount()
        {
            if (_dataTable != null)
            {
                var view = ExcelDataGrid.ItemsSource as DataView;
                int displayedRows = view?.Count ?? _dataTable.Rows.Count;
                int totalRows = _dataTable.Rows.Count;

                if (displayedRows == totalRows)
                {
                    RowCountText.Text = $"{totalRows} rader";
                }
                else
                {
                    RowCountText.Text = $"{displayedRows} av {totalRows} rader";
                }
            }
            else
            {
                RowCountText.Text = "0 rader";
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

            var optsWin = new RisingNumbersOptionsWindow(_risingNumberExceptions, ex =>
            {
                _risingNumberExceptions = ex ?? new List<string>();
            });
            optsWin.Owner = this;
            optsWin.ShowDialog();
            SaveRisingExceptions();

            SaveState();
            var clearedCells = new List<string>();
            for (int ri = 0; ri < _dataTable.Rows.Count; ri++)
            {
                DataRow row = _dataTable.Rows[ri];
                string col5Value = row[4]?.ToString();
                string col6Value = row[5]?.ToString();

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

            var window = new CustomGroupWindow(_customGroups, groupDefinitions =>
            {
                _customGroups = groupDefinitions;
                SaveCustomGroups();
                ApplyCustomSort();
            });
            window.Owner = this;
            window.Show();
        }

        private void CustomCrossSectionButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            if (_customCrossWindow != null && _customCrossWindow.IsLoaded)
            {
                _customCrossWindow.Activate();
                return;
            }

            _customCrossWindow = new CustomCrossSectionWindow(_customCrossRows, rows =>
            {
                _customCrossRows = rows ?? new List<CustomCrossSectionWindow.CrossRow>();
                ApplyCustomCrossSections();
            });
            _customCrossWindow.Owner = this;
            _customCrossWindow.Closed += (_, __) => _customCrossWindow = null;
            _customCrossWindow.Show();
        }

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

                    bool matchDirect = (!has5 || (col5.IndexOf(def.Col5Text, StringComparison.OrdinalIgnoreCase) >= 0)) &&
                                       (!has6 || (col6.IndexOf(def.Col6Text, StringComparison.OrdinalIgnoreCase) >= 0));

                    bool matchSwapped = (!has5 || (col6.IndexOf(def.Col5Text, StringComparison.OrdinalIgnoreCase) >= 0)) &&
                                        (!has6 || (col5.IndexOf(def.Col6Text, StringComparison.OrdinalIgnoreCase) >= 0));

                    if (matchDirect || matchSwapped)
                    {
                        var (c2, c3, c4) = MapCrossOption(def.SelectedOption);
                        row[1] = c2;
                        row[2] = c3;
                        row[3] = c4;
                        _rowsWithCustomCross.Add(i);
                        break;
                    }
                }
            }

            UpdateGrid("Anvendt egendefinert tverrsnitt");
        }

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

        private void SaveRisingExceptions()
        {
            try
            {
                var dir = Path.GetDirectoryName(_risingExceptionsFilePath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(_risingNumberExceptions ?? new List<string>(), options);
                File.WriteAllText(_risingExceptionsFilePath, json);
            }
            catch
            {
            }
        }

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

            var window = new ConvertToDurapartWindow(_conversionRules, rules =>
            {
                _conversionRules = rules;
                SaveConversionRules();
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

            var window = new CustomTextFillWindow(_textFillPatterns, patterns =>
            {
                _textFillPatterns = patterns;
                SaveTextFillPatterns();
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

            var window = new StarDupesWindow(_starDupesRules, rules =>
            {
                _starDupesRules = rules;
                SaveStarDupesRules();
                ApplyStarDupes();
            });
            window.Owner = this;
            window.Show();
        }

        private void RemoveRelayButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            var window = new RemoveRelayWindow(_removeRelayPatterns, (allPatterns, enabledPatterns) =>
            {
                _removeRelayPatterns = allPatterns;
                SaveRemoveRelayPatterns();
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
                        break;
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

            if (_dataTable.Columns.Count >= 6)
            {
                for (int i = 0; i < _dataTable.Rows.Count; i++)
                {
                    string col5Value = _dataTable.Rows[i][4]?.ToString() ?? "";
                    string col6Value = _dataTable.Rows[i][5]?.ToString() ?? "";

                    int col5Priority = GetTextPriority(col5Value);
                    int col6Priority = GetTextPriority(col6Value);

                    bool shouldSwap = false;

                    if (col6Priority > 0 && col5Priority == 0)
                    {
                        shouldSwap = true;
                    }
                    else if (col6Priority > 0 && col5Priority > 0 && col6Priority < col5Priority)
                    {
                        shouldSwap = true;
                    }

                    if (shouldSwap)
                    {
                        _dataTable.Rows[i][4] = col6Value;
                        _dataTable.Rows[i][5] = col5Value;
                    }
                }
            }

            var groupedRows = new List<(GroupDefinition Group, List<DataRow> Rows)>();
            var ungroupedRows = new List<DataRow>();

            var sortedGroups = _customGroups.OrderBy(g => g.Priority).ToList();

            var assignedRows = new HashSet<DataRow>();

            foreach (var group in sortedGroups)
            {
                var matchingRows = new List<DataRow>();

                for (int i = 0; i < _dataTable.Rows.Count; i++)
                {
                    var row = _dataTable.Rows[i];

                    if (assignedRows.Contains(row)) continue;

                    string col5Value = row[4]?.ToString() ?? "";
                    string col6Value = row[5]?.ToString() ?? "";

                    if (col5Value.Contains(group.ContainsText) || col6Value.Contains(group.ContainsText))
                    {
                        matchingRows.Add(row);
                        assignedRows.Add(row);
                    }
                }

                if (matchingRows.Count > 0)
                {
                    matchingRows.Sort((a, b) =>
                    {
                        string aValue = a[4]?.ToString() ?? "";
                        string bValue = b[4]?.ToString() ?? "";
                        return string.Compare(aValue, bValue, StringComparison.Ordinal);
                    });

                    groupedRows.Add((group, matchingRows));
                }
            }

            foreach (DataRow row in _dataTable.Rows)
            {
                if (!assignedRows.Contains(row))
                {
                    ungroupedRows.Add(row);
                }
            }

            DataTable sortedTable = _dataTable.Clone();

            int groupIndex = 1;
            foreach (var (group, rows) in groupedRows)
            {
                foreach (var row in rows)
                {
                    var newRow = sortedTable.NewRow();
                    newRow.ItemArray = row.ItemArray;

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
                groupIndex++;
            }

            foreach (var row in ungroupedRows)
            {
                var newRow = sortedTable.NewRow();
                newRow.ItemArray = row.ItemArray;

                if (sortedTable.Columns.Count > 6)
                {
                    newRow[6] = "";
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

            var cellValues = new Dictionary<string, List<(int rowIndex, int colIndex, string adjacentValue)>>();

            for (int i = 0; i < _dataTable.Rows.Count; i++)
            {
                string col5Value = _dataTable.Rows[i][4]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(col5Value) && !col5Value.EndsWith("*"))
                {
                    string col6Adjacent = _dataTable.Rows[i][5]?.ToString()?.Trim() ?? "";
                    if (!cellValues.ContainsKey(col5Value))
                        cellValues[col5Value] = new List<(int, int, string)>();
                    cellValues[col5Value].Add((i, 4, col6Adjacent));
                }

                string col6Value = _dataTable.Rows[i][5]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(col6Value) && !col6Value.EndsWith("*"))
                {
                    string col5Adjacent = _dataTable.Rows[i][4]?.ToString()?.Trim() ?? "";
                    if (!cellValues.ContainsKey(col6Value))
                        cellValues[col6Value] = new List<(int, int, string)>();
                    cellValues[col6Value].Add((i, 5, col5Adjacent));
                }
            }

            foreach (var kvp in cellValues.Where(x => x.Value.Count > 1))
            {
                string duplicateValue = kvp.Key;
                var locations = kvp.Value;

                (int rowIndex, int colIndex, string adjacentValue)? bestLocation = null;
                int bestPriority = int.MaxValue;

                foreach (var location in locations)
                {
                    foreach (var rule in _starDupesRules.OrderBy(r => r.Priority))
                    {
                        if (duplicateValue.Contains(rule.DuplicateContains) &&
                            location.adjacentValue.Contains(rule.AdjacentContains))
                        {
                            if (rule.Priority < bestPriority)
                            {
                                bestPriority = rule.Priority;
                                bestLocation = location;
                            }
                            break;
                        }
                    }
                }

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

            var patterns = patternsToApply ?? _removeRelayPatterns;

            foreach (DataRow row in _dataTable.Rows)
            {
                for (int j = 0; j < _dataTable.Columns.Count; j++)
                {
                    string cellValue = row[j]?.ToString();
                    if (string.IsNullOrEmpty(cellValue)) continue;

                    foreach (var pattern in patterns)
                    {
                        if (cellValue.Contains(pattern.ContainsText))
                        {
                            modifiedCells.Add($"Rad {_dataTable.Rows.IndexOf(row) + 1}, Kol {j + 1}");
                            row[j] = string.Empty;

                            if (j > 0)
                            {
                                row[j - 1] = string.Empty;
                                modifiedCells.Add($"Rad {_dataTable.Rows.IndexOf(row) + 1}, Kol {j}");
                            }

                            if (j < _dataTable.Columns.Count - 1)
                            {
                                row[j + 1] = string.Empty;
                                modifiedCells.Add($"Rad {_dataTable.Rows.IndexOf(row) + 1}, Kol {j + 2}");
                            }

                            break;
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

            if (_dataTable.Columns.Count < 6)
            {
                UpdateGrid("Trenger minst 6 kolonner for reorganisering");
                return;
            }

            foreach (DataRow row in _dataTable.Rows)
            {
                string col2Original = row[1]?.ToString() ?? "";
                string col3Original = row[2]?.ToString() ?? "";
                string col4Original = row[3]?.ToString() ?? "";
                string col5Original = row[4]?.ToString() ?? "";
                string col6Original = row[5]?.ToString() ?? "";

                row[1] = col5Original;
                row[2] = col6Original;
                row[3] = col2Original;
                row[4] = col3Original;
                row[5] = col4Original;

                reorganizedRows++;
            }

            UpdateGrid($"Reorganisert kolonner (5→2, 6→3, 2→4, 3→5, 4→6) for {reorganizedRows} rader");
        }

        private void ApplyConversionRules()
        {
            if (_dataTable == null) return;

            SaveState();
            var modifiedCells = new List<string>();

            foreach (DataRow row in _dataTable.Rows)
            {
                for (int j = 0; j < _dataTable.Columns.Count; j++)
                {
                    string cellValue = row[j]?.ToString();
                    if (string.IsNullOrEmpty(cellValue)) continue;

                    foreach (var rule in _conversionRules)
                    {
                        if (cellValue.Equals(rule.FromText, StringComparison.Ordinal))
                        {
                            row[j] = rule.ToText;
                            modifiedCells.Add($"Rad {_dataTable.Rows.IndexOf(row) + 1}, Kol {j + 1}: '{rule.FromText}' → '{rule.ToText}'");
                            break;
                        }
                    }
                }
            }

            UpdateGrid($"Konvertert {modifiedCells.Count} celler til Durapart format: {string.Join(", ", modifiedCells.Take(5))}{(modifiedCells.Count > 5 ? "..." : "")}");
        }

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
            return 0;
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

        private void ExcelDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
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
                    _hasUnsavedChanges = false;
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Feil ved lagring av fil - {ex.Message}";
            }
        }

        private void ProcessRawDataButton_Click(object sender, RoutedEventArgs e)
        {
            var processWindow = new ProcessRawDataWindow(filePath =>
            {
                if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                {
                    _filePath = filePath;
                    FilePathText.Text = Path.GetFileName(_filePath);
                    _undoStack.Clear();
                    _rowsRemoved = 0;
                    _removeEqualsApplied = false;
                    _removeLVApplied = false;
                    _hasUnsavedChanges = false;
                    LoadExcelFile();
                }
            });
            processWindow.Owner = this;
            processWindow.Show();
        }

        private void MarkUnmarkedDupesButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            var window = new MarkUnmarkedDupesWindow(patterns =>
            {
                ApplyMarkUnmarkedDupes(patterns);
            });
            window.Owner = this;
            window.ShowDialog();
        }

        private void ApplyMarkUnmarkedDupes(List<string> patterns)
        {
            if (_dataTable == null || patterns.Count == 0) return;

            ExcelDataGrid.Items.Refresh();

            var cellOccurrences = new Dictionary<string, List<(int rowIndex, int colIndex)>>();

            for (int i = 0; i < _dataTable.Rows.Count; i++)
            {
                for (int colIdx = 4; colIdx <= 5; colIdx++)
                {
                    string cellValue = _dataTable.Rows[i][colIdx]?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(cellValue)) continue;

                    bool matchesPattern = patterns.Any(p => cellValue.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (!matchesPattern) continue;

                    string cleanValue = cellValue.TrimEnd('*');

                    if (!cellOccurrences.ContainsKey(cleanValue))
                        cellOccurrences[cleanValue] = new List<(int, int)>();

                    cellOccurrences[cleanValue].Add((i, colIdx));
                }
            }

            int markedCount = 0;
            foreach (var kvp in cellOccurrences.Where(x => x.Value.Count > 1))
            {
                foreach (var (rowIndex, colIndex) in kvp.Value)
                {
                    string cellValue = _dataTable.Rows[rowIndex][colIndex]?.ToString()?.Trim() ?? "";
                    if (!cellValue.EndsWith("*"))
                    {
                        markedCount++;
                    }
                }
            }

            UpdateGrid($"Markerte {markedCount} umerkede duplikater med gul farge");
            MessageBox.Show($"Markerte {markedCount} umerkede duplikater med gul farge i tabellen.",
                "Markering fullført", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void AutoFilterButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Ingen fil lastet";
                return;
            }

            var result = MessageBox.Show(
                "Dette vil automatisk kjøre følgende filtre:\n\n" +
                "1. Fjern =+LV+MX+XS\n" +
                "2. Fjern Duplikater\n" +
                "3-9. Funksjoner med vinduer (stopper for brukerinput)\n\n" +
                "Vil du fortsette?",
                "Auto-Filtrer",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes) return;

            RemoveAllButton_Click(sender, e);
            RemoveDupesButton_Click(sender, e);

            StatusText.Text = "Auto-filtrer: Klar for steg 3 (Stigende Tall). Klikk på knappen for å fortsette.";
        }
    }

    public class GroupConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is string colorIndexStr && int.TryParse(colorIndexStr, out int colorIndex))
            {
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