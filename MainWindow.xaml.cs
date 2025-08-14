using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
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

        public MainWindow()
        {
            InitializeComponent();
            _undoStack = new Stack<DataTable>();
            _rowsRemoved = 0;
            _removeEqualsApplied = false;
            _removeLVApplied = false;
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

            SaveState();

            if (_dataTable.Columns.Count >= 6)
            {
                for (int i = 0; i < _dataTable.Rows.Count; i++)
                {
                    string col5Value = _dataTable.Rows[i][4]?.ToString() ?? "";
                    string col6Value = _dataTable.Rows[i][5]?.ToString() ?? "";
                    bool col5IsTarget = col5Value.Contains("X1:") || col5Value.Contains("F11-X");
                    bool col6IsTarget = col6Value.Contains("X1:") || col6Value.Contains("F11-X");
                    if (col6IsTarget && !col5IsTarget)
                    {
                        _dataTable.Rows[i][4] = col6Value;
                        _dataTable.Rows[i][5] = col5Value;
                    }
                }
            }

            var x1Rows = new List<(string Prefix, DataRow Row)>();
            var f11Rows = new List<(string Prefix, DataRow Row)>();
            var otherRows = new List<DataRow>();
            for (int i = 0; i < _dataTable.Rows.Count; i++)
            {
                string col5Value = _dataTable.Rows[i][4]?.ToString() ?? "";
                if (col5Value.Contains("X1:"))
                {
                    string prefix = col5Value.Split('-').FirstOrDefault() ?? col5Value;
                    x1Rows.Add((prefix, _dataTable.Rows[i]));
                }
                else if (col5Value.Contains("F11-X"))
                {
                    var match = Regex.Match(col5Value, @"^(.*F11-X\d+)(?::|$)");
                    string prefix = match.Success ? match.Groups[1].Value : col5Value;
                    f11Rows.Add((prefix, _dataTable.Rows[i]));
                }
                else
                {
                    otherRows.Add(_dataTable.Rows[i]);
                }
            }

            x1Rows.Sort((a, b) => string.Compare(a.Prefix, b.Prefix, StringComparison.Ordinal));
            f11Rows.Sort((a, b) =>
            {
                int prefixCompare = string.Compare(a.Prefix, b.Prefix, StringComparison.Ordinal);
                if (prefixCompare != 0) return prefixCompare;
                return string.Compare(a.Row[4]?.ToString(), b.Row[4]?.ToString(), StringComparison.Ordinal);
            });

            DataTable sortedTable = _dataTable.Clone();
            int groupNumber = 0;

            string lastX1Prefix = null;
            foreach (var x1Row in x1Rows)
            {
                if (x1Row.Prefix != lastX1Prefix)
                {
                    lastX1Prefix = x1Row.Prefix;
                    groupNumber++;
                }
                var row = sortedTable.NewRow();
                row.ItemArray = x1Row.Row.ItemArray;
                if (sortedTable.Columns.Count > 6)
                {
                    row[6] = groupNumber.ToString();
                }
                sortedTable.Rows.Add(row);
            }

            string lastF11Prefix = null;
            foreach (var f11Row in f11Rows)
            {
                if (f11Row.Prefix != lastF11Prefix)
                {
                    lastF11Prefix = f11Row.Prefix;
                    groupNumber++;
                }
                var row = sortedTable.NewRow();
                row.ItemArray = f11Row.Row.ItemArray;
                if (sortedTable.Columns.Count > 6)
                {
                    row[6] = groupNumber.ToString();
                }
                sortedTable.Rows.Add(row);
            }

            foreach (var otherRow in otherRows)
            {
                sortedTable.Rows.Add(otherRow.ItemArray);
            }

            _dataTable = sortedTable;
            MoveCellsUpward();
            UpdateGrid("Status: Data sorted with X1: and F11-X cells in column 5, group numbers assigned");
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

                    int lastGroupNumber = -1;
                    bool useFirstColor = true;

                    for (int i = 0; i < _dataTable.Rows.Count; i++)
                    {
                        var row = worksheet.Row(i + 2);
                        string groupNumberStr = _dataTable.Rows[i][6]?.ToString() ?? "";
                        if (int.TryParse(groupNumberStr, out int currentGroupNumber))
                        {
                            if (currentGroupNumber != lastGroupNumber)
                            {
                                lastGroupNumber = currentGroupNumber;
                                useFirstColor = !useFirstColor;
                            }
                            XLColor rowColor = useFirstColor ? XLColor.LightGreen : XLColor.LightBlue;
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
            if (value is string groupNumberStr && int.TryParse(groupNumberStr, out int groupNumber))
            {
                return groupNumber % 2 == 1 ? Brushes.LightGreen : Brushes.LightBlue;
            }
            return Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}