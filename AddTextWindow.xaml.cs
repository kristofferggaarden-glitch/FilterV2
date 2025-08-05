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
    public partial class AddTextWindow : Window
    {
        public class CellPair
        {
            public string FirstCell { get; set; }
            public string SecondCell { get; set; }
            public string Ledningstype { get; set; }
            public string HylseSide1 { get; set; }
            public string HylseSide2 { get; set; }
        }

        private readonly Action<List<CellPair>> _onApply;
        private readonly List<CellPair> _cellPairs;
        private readonly string _jsonFilePath;

        public AddTextWindow(Action<List<CellPair>> onApply)
        {
            InitializeComponent();
            _onApply = onApply;
            _cellPairs = new List<CellPair>();
            _jsonFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FilterV1", "add_text.json");
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
                        _cellPairs.AddRange(loadedPairs.Distinct(new CellPairComparer()));
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

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            string pastedText = PasteTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(pastedText) || pastedText == "Paste cell pairs (e.g., value1\tvalue2)...")
            {
                MessageBox.Show("Please paste or enter valid cell pairs.", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string[] parts = pastedText.Split('\t', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2)
            {
                string firstCell = parts[0].Trim();
                string secondCell = parts[1].Trim();
                string ledningstype = parts.Length > 2 ? parts[2].Trim() : "";
                string hylseSide1 = parts.Length > 3 ? parts[3].Trim() : "";
                string hylseSide2 = parts.Length > 4 ? parts[4].Trim() : "";
                if (!string.IsNullOrWhiteSpace(firstCell) && !string.IsNullOrWhiteSpace(secondCell))
                {
                    var newPair = new CellPair
                    {
                        FirstCell = firstCell,
                        SecondCell = secondCell,
                        Ledningstype = ledningstype,
                        HylseSide1 = hylseSide1,
                        HylseSide2 = hylseSide2
                    };
                    if (!_cellPairs.Any(p => new CellPairComparer().Equals(p, newPair)))
                    {
                        _cellPairs.Add(newPair);
                        CellPairsGrid.Items.Refresh();
                        PasteTextBox.Text = string.Empty;
                        PasteTextBox.Focus();
                        SaveCellPairs();
                    }
                    else
                    {
                        MessageBox.Show("This pair already exists.", "Duplicate Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("First Cell and Second Cell must be non-empty.", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                MessageBox.Show("Pasted text must contain at least two values separated by tabs.", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                    string[] parts = line.Split('\t', StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 2)
                    {
                        string firstCell = parts[0].Trim();
                        string secondCell = parts[1].Trim();
                        string ledningstype = parts.Length > 2 ? parts[2].Trim() : "";
                        string hylseSide1 = parts.Length > 3 ? parts[3].Trim() : "";
                        string hylseSide2 = parts.Length > 4 ? parts[4].Trim() : "";
                        if (!string.IsNullOrWhiteSpace(firstCell) && !string.IsNullOrWhiteSpace(secondCell))
                        {
                            var newPair = new CellPair
                            {
                                FirstCell = firstCell,
                                SecondCell = secondCell,
                                Ledningstype = ledningstype,
                                HylseSide1 = hylseSide1,
                                HylseSide2 = hylseSide2
                            };
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
                }

                e.Handled = true;
            }
        }

        private void CellPairsGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                SaveCellPairs();
            }
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = CellPairsGrid.SelectedItems.Cast<CellPair>().ToList();
            if (selectedItems.Any())
            {
                foreach (var item in selectedItems)
                {
                    _cellPairs.Remove(item);
                }
                CellPairsGrid.Items.Refresh();
                SaveCellPairs();
                MessageBox.Show($"Removed {selectedItems.Count} pair(s).", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Please select at least one pair to remove.", "Selection Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            _onApply(_cellPairs);
            SaveCellPairs();
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void AutoFillButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedOption = (AutoFillComboBox.SelectedItem as ComboBoxItem)?.Content.ToString();
            if (selectedOption == "Custom" || string.IsNullOrEmpty(selectedOption))
                return;

            string ledningstype = "";
            string hylseSide1 = "";
            string hylseSide2 = "";

            switch (selectedOption)
            {
                case "UNIBK1.0":
                    ledningstype = "UNIBK1.0";
                    hylseSide1 = "HYLSE 1.0";
                    hylseSide2 = "HYLSE 1.0";
                    break;
                case "UNIBK1.5":
                    ledningstype = "UNIBK1.5";
                    hylseSide1 = "HYLSE 1.5";
                    hylseSide2 = "HYLSE 1.5";
                    break;
                case "UNIBK2.5":
                    ledningstype = "UNIBK2.5";
                    hylseSide1 = "HYLSE 2.5";
                    hylseSide2 = "HYLSE 2.5";
                    break;
                case "UNIBK4.0":
                    ledningstype = "UNIBK4.0";
                    hylseSide1 = "HYLSE 4.0";
                    hylseSide2 = "HYLSE 4.0";
                    break;
            }

            foreach (var pair in _cellPairs)
            {
                pair.Ledningstype = ledningstype;
                pair.HylseSide1 = hylseSide1;
                pair.HylseSide2 = hylseSide2;
            }

            CellPairsGrid.Items.Refresh();
            SaveCellPairs();
        }

        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox != null && textBox.Text == "Paste cell pairs (e.g., value1\tvalue2)...")
            {
                textBox.Text = string.Empty;
            }
        }

        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox != null && string.IsNullOrWhiteSpace(textBox.Text))
            {
                textBox.Text = "Paste cell pairs (e.g., value1\tvalue2)...";
            }
        }
    }
}