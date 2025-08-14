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
        private readonly string _settingsFilePath;

        public MainWindow()
        {
            InitializeComponent();
            _undoStack = new Stack<DataTable>();
            _rowsRemoved = 0;
            _removeEqualsApplied = false;
            _removeLVApplied = false;

            // Set up settings file path
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appFolder = Path.Combine(appDataPath, "FilterV1");
            Directory.CreateDirectory(appFolder);
            _settingsFilePath = Path.Combine(appFolder, "GroupSettings.json");

            // Load custom groups from settings or use defaults
            LoadCustomGroups();
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

        private List<GroupDefinition> GetDefaultGroups()
        {
            return new List<GroupDefinition>
            {
                new GroupDefinition { GroupName = "X1 Group", ContainsText = "X1:", Priority = 1 },
                new GroupDefinition { GroupName = "F11 Group", ContainsText = "F11-X", Priority = 2 }
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
            window.Show(); // Changed to non-modal
        }

        private void AddTextButton_Click(object sender, RoutedEventArgs e)
        {
            var window = new AddTextWindow(pairs =>
            {
                if (_dataTable == null)
                {
                    UpdateGrid("Status: No file loaded, pairs defined but not applied");
                    return;
                }

                SaveState();
                var modifiedRows = new List<int>();
                foreach (DataRow row in _dataTable.Rows)
                {
                    string col5Value = row[4]?.ToString()?.Trim();
                    string col6Value = row[5]?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(col5Value) || string.IsNullOrEmpty(col6Value))
                        continue;

                    foreach (var pair in pairs)
                    {
                        if (string.Equals(col5Value, pair.FirstCell, StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(col6Value, pair.SecondCell, StringComparison.OrdinalIgnoreCase))
                        {
                            int rowIndex = _dataTable.Rows.IndexOf(row) + 1;
                            row[1] = pair.Ledningstype ?? "";
                            row[2] = pair.HylseSide1 ?? "";
                            row[3] = pair.HylseSide2 ?? "";
                            modifiedRows.Add(rowIndex);
                            break;
                        }
                    }
                }
                UpdateGrid($"Status: Added text to columns 2-4 in rows: {string.Join(", ", modifiedRows.OrderBy(x => x))}");
            });
            window.Show(); // Changed to non-modal
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

        private void ApplyCustomSort()
        {
            if (_dataTable == null) return;

            SaveState();

            // Swap columns 5 and 6 if column 6 contains target text and column 5 doesn't
            if (_dataTable.Columns.Count >= 6)
            {
                for (int i = 0; i < _dataTable.Rows.Count; i++)
                {
                    string col5Value = _dataTable.Rows[i][4]?.ToString() ?? "";
                    string col6Value = _dataTable.Rows[i][5]?.ToString() ?? "";

                    bool col5HasTarget = _customGroups.Any(g => col5Value.Contains(g.ContainsText) || col6Value.Contains(g.ContainsText));
                    bool col6HasTarget = _customGroups.Any(g => col6Value.Contains(g.ContainsText) || col5Value.Contains(g.ContainsText));

                    if (col6HasTarget && !col5HasTarget)
                    {
                        _dataTable.Rows[i][4] = col6Value;
                        _dataTable.Rows[i][5] = col5Value;
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

            var activeGroups = groupedRows.Select(g => g.Group.GroupName).ToList();
            string groupSummary = string.Join(", ", activeGroups);
            UpdateGrid($"Status: Data sorted with custom groups: {groupSummary}");
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