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
        private readonly string _jsonFilePath;

        public RemoveCellsWindow(Action<List<CellPair>> onApply)
        {
            InitializeComponent();
            _onApply = onApply;
            _cellPairs = new List<CellPair>();
            _jsonFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FilterV1", "remove_cells.json");
            LoadCellPairs();
            CellPairsGrid.ItemsSource = _cellPairs;
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
                            _cellPairs.Add(newPair);
                            addedCount++;
                        }
                    }
                }
            }

            if (addedCount > 0)
            {
                CellPairsGrid.Items.Refresh();
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
                    _cellPairs.AddRange(newPairs);
                    CellPairsGrid.Items.Refresh();
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
                string message = selectedItems.Count == 1 ?
                    "Are you sure you want to remove the selected pattern pair?" :
                    $"Are you sure you want to remove {selectedItems.Count} selected pattern pairs?";

                if (MessageBox.Show(message, "Confirm Removal", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    foreach (var item in selectedItems)
                    {
                        _cellPairs.Remove(item);
                    }
                    CellPairsGrid.Items.Refresh();
                    SaveCellPairs();
                    MessageBox.Show($"Removed {selectedItems.Count} pattern pair(s).", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
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
    }
}