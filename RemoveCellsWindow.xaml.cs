using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FilterV1
{
    public partial class RemoveCellsWindow : Window
    {
        public class CellPair
        {
            public string FirstCell { get; set; } = "";
            public string SecondCell { get; set; } = "";
        }

        private readonly Action<List<CellPair>> _onApply;
        private readonly List<CellPair> _cellPairs;
        // Holds the subset of cell pairs after search filtering. When no filter is applied,
        // this list contains all pairs. Binding the DataGrid to this list allows dynamic
        // updates when the user types in the search box.
        private List<CellPair> _filteredCellPairs;
        private readonly string _jsonFilePath;

        public RemoveCellsWindow(Action<List<CellPair>> onApply)
        {
            InitializeComponent();
            _onApply = onApply;
            _cellPairs = new List<CellPair>();
            _filteredCellPairs = new List<CellPair>();
            _jsonFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FilterV1", "remove_cells.json");
            LoadCellPairs();
            // Bind the DataGrid to the filtered collection. When filters are applied the
            // contents of this list are replaced and the grid refreshed.
            CellPairsGrid.ItemsSource = _filteredCellPairs;
        }

        private void LoadCellPairs()
        {
            try
            {
                if (File.Exists(_jsonFilePath))
                {
                    string json = File.ReadAllText(_jsonFilePath);
                    var loadedPairs = JsonSerializer.Deserialize<List<CellPair>>(json);
                    if (loadedPairs != null)
                    {
                        // Filter out null or empty pairs and remove duplicates
                        var validPairs = loadedPairs
                            .Where(p => p != null && !string.IsNullOrWhiteSpace(p.FirstCell) && !string.IsNullOrWhiteSpace(p.SecondCell))
                            .Distinct(new CellPairComparer());
                        _cellPairs.AddRange(validPairs);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading cell pairs: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            // Populate the filtered list with all pairs initially
            _filteredCellPairs.Clear();
            _filteredCellPairs.AddRange(_cellPairs);
            CellPairsGrid.Items.Refresh();
        }

        private void SaveCellPairs()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_jsonFilePath));
                string json = JsonSerializer.Serialize(_cellPairs, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_jsonFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving cell pairs: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PasteTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddButton_Click(sender, e);
            }
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            string pastedText = PasteTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(pastedText) || pastedText == "Paste cell pairs here...")
            {
                MessageBox.Show("Please paste valid cell pairs or enter text patterns.", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Handle multi-line paste - split by newlines and process each line
            string[] lines = pastedText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int addedCount = 0;

            foreach (string line in lines)
            {
                string trimmedLine = line.Trim();
                if (string.IsNullOrEmpty(trimmedLine)) continue;

                // Split by tab (Excel paste) or space if no tab
                string[] parts = trimmedLine.Contains('\t') ?
                    trimmedLine.Split('\t', StringSplitOptions.RemoveEmptyEntries) :
                    trimmedLine.Split(' ', StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length >= 2)
                {
                    string firstCell = parts[0].Trim();
                    string secondCell = parts[1].Trim();

                    if (!string.IsNullOrWhiteSpace(firstCell) && !string.IsNullOrWhiteSpace(secondCell))
                    {
                        var newPair = new CellPair { FirstCell = firstCell, SecondCell = secondCell };
                        if (!_cellPairs.Any(p => new CellPairComparer().Equals(p, newPair)))
                        {
                            // Add new pair to the underlying collection
                            _cellPairs.Add(newPair);
                            addedCount++;
                        }
                    }
                }
            }

            if (addedCount > 0)
            {
                // Update the filtered list based on the current search term so that
                // newly added pairs appear immediately if they match the filter.
                FilterCellPairs(SearchTextBox?.Text ?? string.Empty);
                SaveCellPairs();
                PasteTextBox.Text = string.Empty;
                PasteTextBox.Focus();

                // Suppress informational dialog when adding multiple pairs
            }
            else
            {
                MessageBox.Show("No valid pairs found or all pairs already exist.", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void PasteTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                string clipboardText = Clipboard.GetText();
                if (string.IsNullOrEmpty(clipboardText))
                {
                    MessageBox.Show("No text in clipboard.", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    e.Handled = true;
                    return;
                }

                bool addedNew = false;
                var newPairs = new List<CellPair>();
                string[] lines = clipboardText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string line in lines)
                {
                    string trimmedLine = line.Trim();
                    if (string.IsNullOrEmpty(trimmedLine)) continue;

                    // Split by tab (Excel paste) or space if no tab  
                    string[] parts = trimmedLine.Contains('\t') ?
                        trimmedLine.Split('\t', StringSplitOptions.RemoveEmptyEntries) :
                        trimmedLine.Split(' ', StringSplitOptions.RemoveEmptyEntries);

                    if (parts.Length >= 2)
                    {
                        string firstCell = parts[0].Trim();
                        string secondCell = parts[1].Trim();

                        if (!string.IsNullOrWhiteSpace(firstCell) && !string.IsNullOrWhiteSpace(secondCell))
                        {
                            var newPair = new CellPair { FirstCell = firstCell, SecondCell = secondCell };
                            if (!_cellPairs.Any(p => new CellPairComparer().Equals(p, newPair)))
                            {
                                newPairs.Add(newPair);
                                addedNew = true;
                            }
                        }
                    }
                }

                if (addedNew)
                {
                    // Add to underlying list then refresh filtered view
                    _cellPairs.AddRange(newPairs);
                    FilterCellPairs(SearchTextBox?.Text ?? string.Empty);
                    SaveCellPairs();
                    PasteTextBox.Text = string.Empty;
                    PasteTextBox.Focus();

                    // Suppress informational dialog when adding multiple pairs
                }

                e.Handled = true;
            }
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = CellPairsGrid.SelectedItems.Cast<CellPair>().ToList();
            if (selectedItems.Any())
            {
                // Immediately remove the selected pairs without confirmation or success prompts.
                foreach (var item in selectedItems)
                {
                    _cellPairs.Remove(item);
                }
                FilterCellPairs(SearchTextBox?.Text ?? string.Empty);
                SaveCellPairs();
            }
            else
            {
                MessageBox.Show("Please select at least one pattern pair to remove.", "Selection Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_cellPairs.Count == 0)
            {
                MessageBox.Show("Please add at least one pattern pair before applying.", "No Patterns", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _onApply(_cellPairs);
            SaveCellPairs();
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox != null && textBox.Text == "Paste cell pairs here...")
            {
                textBox.Text = string.Empty;
            }
        }

        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox != null && string.IsNullOrWhiteSpace(textBox.Text))
            {
                textBox.Text = "Paste cell pairs here...";
            }
        }

        /// <summary>
        /// Invoked whenever the search text changes. Filters the list of pattern pairs
        /// based on the provided search term. The search is case-insensitive and checks
        /// both the first and second cell values. When the search term is empty,
        /// all pairs are displayed.
        /// </summary>
        /// <param name="sender">The search TextBox.</param>
        /// <param name="e">Event arguments.</param>
        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string searchTerm = (sender as TextBox)?.Text ?? string.Empty;
            FilterCellPairs(searchTerm);
        }

        /// <summary>
        /// Filters the underlying collection of cell pairs (_cellPairs) and populates
        /// _filteredCellPairs with items that contain the search term. Matching is
        /// case-insensitive and checks both columns. After filtering, the DataGrid is
        /// refreshed to reflect the changes.
        /// </summary>
        /// <param name="searchTerm">The case-insensitive search string.</param>
        private void FilterCellPairs(string searchTerm)
        {
            _filteredCellPairs.Clear();
            if (string.IsNullOrWhiteSpace(searchTerm))
            {
                _filteredCellPairs.AddRange(_cellPairs);
            }
            else
            {
                string term = searchTerm.Trim();
                foreach (var pair in _cellPairs)
                {
                    if (pair.FirstCell?.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        pair.SecondCell?.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        _filteredCellPairs.Add(pair);
                    }
                }
            }
            CellPairsGrid.Items.Refresh();
        }
    }
}