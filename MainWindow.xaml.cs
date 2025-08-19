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
using System.Windows.Media;

namespace FilterV1
{
    public partial class MainWindow : Window
    {
        private string _filePath;
        private DataTable _dataTable;
        private Stack<DataTable> _undoStack;
        private int _rowsRemoved;
        private bool _removeEqualsApplied;
        private bool _removeLVApplied;
        private List<GroupDefinition> _customGroups;
        private List<TextFillPattern> _textFillPatterns;
        private List<StarDupesRule> _starDupesRules;
        private readonly string _settingsFilePath;
        private readonly string _textFillSettingsFilePath;
        private readonly string _starDupesSettingsFilePath;

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

            // Load custom groups, text fill patterns, and star dupes rules from settings
            LoadCustomGroups();
            LoadTextFillPatterns();
            LoadStarDupesRules();
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
                FilePathText.Text = _filePath;
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

                    ExcelDataGrid.ItemsSource = _dataTable.DefaultView;
                    UpdateGrid($"Status: File loaded successfully. Total rows: {_dataTable.Rows.Count}");
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Status: Error loading file - {ex.Message}";
            }
        }

        private void SaveState()
        {
            DataTable clone = _dataTable.Copy();
            _undoStack.Push(clone);
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
            ExcelDataGrid.ItemsSource = _dataTable.DefaultView;
            StatusText.Text = $"{statusMessage} (Rows removed: {_rowsRemoved}, Total rows: {_dataTable.Rows.Count})";
        }

        private void RemoveEqualsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Status: No file loaded";
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
            MoveCellsUpward();
            _removeEqualsApplied = true;
            UpdateGrid("Status: '=' removed from cells");
        }

        private void RemoveLVButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Status: No file loaded";
                return;
            }

            SaveState();
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
            MoveCellsUpward();
            _removeLVApplied = true;
            UpdateGrid("Status: '+LV' removed from cells");
        }

        private void RemoveXSButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Status: No file loaded";
                return;
            }

            SaveState();
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
            UpdateGrid("Status: Cells containing 'XS' and adjacent cells cleared");
        }

        private void RemoveMXButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Status: No file loaded";
                return;
            }

            SaveState();
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
            MoveCellsUpward();
            UpdateGrid("Status: Cells containing 'MX' and adjacent cells cleared");
        }

        private void RemoveRisingNumbersButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Status: No file loaded";
                return;
            }

            SaveState();
            foreach (DataRow row in _dataTable.Rows)
            {
                for (int j = 0; j < _dataTable.Columns.Count - 1; j++)
                {
                    string leftCell = row[j]?.ToString();
                    string rightCell = row[j + 1]?.ToString();
                    if (string.IsNullOrEmpty(leftCell) || string.IsNullOrEmpty(rightCell))
                        continue;

                    var matchLeft = Regex.Match(leftCell, @"^(.*):(\d+)$");
                    var matchRight = Regex.Match(rightCell, @"^(.*):(\d+)$");
                    if (matchLeft.Success && matchRight.Success)
                    {
                        string leftPrefix = matchLeft.Groups[1].Value;
                        string rightPrefix = matchRight.Groups[1].Value;
                        if (leftPrefix == rightPrefix)
                        {
                            if (int.TryParse(matchLeft.Groups[2].Value, out int leftNumber) &&
                                int.TryParse(matchRight.Groups[2].Value, out int rightNumber))
                            {
                                if (rightNumber == leftNumber + 1)
                                {
                                    row[j] = string.Empty;
                                    row[j + 1] = string.Empty;
                                }
                            }
                        }
                    }
                }
            }
            MoveCellsUpward();
            UpdateGrid("Status: Rising number pairs cleared");
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
                            if (string.Equals(col5Value, pair.FirstCell, StringComparison.OrdinalIgnoreCase) &&
                                string.Equals(col6Value, pair.SecondCell, StringComparison.OrdinalIgnoreCase))
                            {
                                matchedRows.Add(i + 1);
                                _dataTable.Rows.RemoveAt(i);
                                break;
                            }
                        }
                    }
                    MoveCellsUpward();
                    UpdateGrid($"Status: Removed rows with matching cells in columns 5-6: {string.Join(", ", matchedRows.OrderByDescending(x => x))}");
                }
                else
                {
                    UpdateGrid("Status: No file loaded, pairs defined but not applied");
                }
            });
            window.Show();
        }

        private void RemoveDupesButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Status: No file loaded";
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
            UpdateGrid("Status: Duplicate adjacent cells cleared");
        }

        private void SortButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Status: No file loaded";
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
            window.ShowDialog();
        }

        private void CustomTextFillButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null)
            {
                StatusText.Text = "Status: No file loaded";
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
                StatusText.Text = "Status: No file loaded";
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
            window.ShowDialog();
        }

        private void ApplyTextFillPatterns()
        {
            if (_dataTable == null) return;

            SaveState();
            var modifiedRows = new List<int>();

            foreach (DataRow row in _dataTable.Rows)
            {
                string col5Value = row[4]?.ToString()?.Trim() ?? "";
                string col6Value = row[5]?.ToString()?.Trim() ?? "";

                // Check both columns 5 and 6 for matching patterns in priority order
                foreach (var pattern in _textFillPatterns.OrderBy(p => p.Priority))
                {
                    if (col5Value.Contains(pattern.ContainsText) || col6Value.Contains(pattern.ContainsText))
                    {
                        int rowIndex = _dataTable.Rows.IndexOf(row) + 1;
                        var (col2, col3, col4) = GetOptionValues(pattern.SelectedOption);

                        row[1] = col2; // Column 2
                        row[2] = col3; // Column 3
                        row[3] = col4; // Column 4

                        modifiedRows.Add(rowIndex);
                        break; // Only apply the first matching pattern (highest priority)
                    }
                }
            }

            UpdateGrid($"Status: Applied text fill to columns 2-4 in rows: {string.Join(", ", modifiedRows.OrderBy(x => x))}");
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

            // Add grouped rows with sequential color indices (0, 1, 0, 1, ...)
            int colorIndex = 0;
            foreach (var (group, rows) in groupedRows)
            {
                foreach (var row in rows)
                {
                    var newRow = sortedTable.NewRow();
                    newRow.ItemArray = row.ItemArray;

                    // Set color index for alternating colors
                    if (sortedTable.Columns.Count > 6)
                    {
                        newRow[6] = colorIndex.ToString();
                    }
                    else if (sortedTable.Columns.Count == 6)
                    {
                        sortedTable.Columns.Add("Group", typeof(string));
                        newRow[6] = colorIndex.ToString();
                    }

                    sortedTable.Rows.Add(newRow);
                }
                colorIndex = (colorIndex + 1) % 2; // Alternate between 0 and 1
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
            UpdateGrid($"Status: Data sorted with priority text moved to left. Groups: {groupSummary}");
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
                    modifiedCells.Add($"Row {bestLocation.Value.rowIndex + 1}, Column {bestLocation.Value.colIndex + 1}");
                }
            }

            UpdateGrid($"Status: Applied star marking to {modifiedCells.Count} duplicate cells: {string.Join(", ", modifiedCells.Take(10))}{(modifiedCells.Count > 10 ? "..." : "")}");
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
                StatusText.Text = "Status: No actions to undo";
                return;
            }

            _dataTable = _undoStack.Pop();
            _rowsRemoved = Math.Max(0, _rowsRemoved - (_dataTable.Rows.Count - _undoStack.Count));
            _removeEqualsApplied = _undoStack.Any() && _undoStack.Peek().Rows.Cast<DataRow>().Any(r => r.ItemArray.Any(v => v?.ToString().Contains("=") == false));
            _removeLVApplied = _undoStack.Any() && _undoStack.Peek().Rows.Cast<DataRow>().Any(r => r.ItemArray.Any(v => v?.ToString().Contains("+LV") == false));
            UpdateGrid("Status: Last action undone");
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null || string.IsNullOrEmpty(_filePath))
            {
                StatusText.Text = "Status: No file loaded";
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
                            // Alternate colors: 0 = green, 1 = pink
                            XLColor rowColor = colorIndex == 0 ? XLColor.LightGreen : XLColor.LightPink;
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
                    StatusText.Text = "Status: File overwritten successfully with colors";
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Status: Error saving file - {ex.Message}";
            }
        }
    }

    public class GroupConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is string colorIndexStr && int.TryParse(colorIndexStr, out int colorIndex))
            {
                // Alternate colors: 0 = green, 1 = pink
                return colorIndex == 0 ? Brushes.LightGreen : Brushes.LightPink;
            }
            return Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}