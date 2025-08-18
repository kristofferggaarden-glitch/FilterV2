using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FilterV1
{
    public partial class CustomTextFillWindow : Window
    {
        private Action<List<TextFillPattern>> _callback;
        private List<TextFillPattern> _textFillPatterns;

        public CustomTextFillWindow(List<TextFillPattern> existingPatterns, Action<List<TextFillPattern>> callback)
        {
            InitializeComponent();
            _callback = callback;

            // Create a deep copy of existing patterns
            _textFillPatterns = existingPatterns.Select(p => new TextFillPattern
            {
                ContainsText = p.ContainsText,
                SelectedOption = p.SelectedOption
            }).ToList();

            OptionComboBox.SelectionChanged += OptionComboBox_SelectionChanged;
            RefreshPatternsList();
            UpdatePreview();
        }

        private void RefreshPatternsList()
        {
            PatternsListBox.Items.Clear();
            foreach (var pattern in _textFillPatterns)
            {
                string optionText = GetOptionDescription(pattern.SelectedOption);
                PatternsListBox.Items.Add($"Text: '{pattern.ContainsText}' → {optionText}");
            }
        }

        private string GetOptionDescription(int option)
        {
            switch (option)
            {
                case 1: return "Option 1 (UNIBK1.0, HYLSE 1.0, HYLSE 1.0)";
                case 2: return "Option 2 (UNIBK1.5, HYLSE 1.5, HYLSE 1.5)";
                case 3: return "Option 3 (UNIBK2.5, HYLSE 2.5, HYLSE 2.5)";
                case 4: return "Option 4 (UNIBK4.0, HYLSE 4.0, HYLSE 4.0)";
                default: return "Option 1 (UNIBK1.0, HYLSE 1.0, HYLSE 1.0)";
            }
        }

        private void UpdatePreview()
        {
            var selectedItem = OptionComboBox.SelectedItem as ComboBoxItem;
            if (selectedItem != null)
            {
                int option = int.Parse(selectedItem.Tag.ToString());
                var (col2, col3, col4) = GetOptionValues(option);

                PreviewCol2.Text = col2;
                PreviewCol3.Text = col3;
                PreviewCol4.Text = col4;
            }
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

        private void ContainsTextTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddPatternButton_Click(sender, e);
            }
        }

        private void OptionComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void AddPatternButton_Click(object sender, RoutedEventArgs e)
        {
            string containsText = ContainsTextTextBox.Text.Trim();

            if (string.IsNullOrEmpty(containsText))
            {
                MessageBox.Show("Please enter the text pattern to match.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Check for duplicates
            if (_textFillPatterns.Any(p => p.ContainsText.Equals(containsText, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("A pattern with this text already exists.", "Duplicate Entry",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var selectedItem = OptionComboBox.SelectedItem as ComboBoxItem;
            int selectedOption = selectedItem != null ? int.Parse(selectedItem.Tag.ToString()) : 1;

            _textFillPatterns.Add(new TextFillPattern
            {
                ContainsText = containsText,
                SelectedOption = selectedOption
            });

            ContainsTextTextBox.Clear();
            ContainsTextTextBox.Focus(); // Keep focus for easy multiple entries
            RefreshPatternsList();
        }

        private void EditPatternButton_Click(object sender, RoutedEventArgs e)
        {
            if (PatternsListBox.SelectedIndex >= 0)
            {
                var selectedPattern = _textFillPatterns[PatternsListBox.SelectedIndex];

                // Load the selected pattern into the input fields
                ContainsTextTextBox.Text = selectedPattern.ContainsText;
                OptionComboBox.SelectedIndex = selectedPattern.SelectedOption - 1; // Convert to 0-based index

                // Remove the old pattern (it will be re-added when user clicks Add)
                _textFillPatterns.RemoveAt(PatternsListBox.SelectedIndex);
                RefreshPatternsList();

                ContainsTextTextBox.Focus();
            }
            else
            {
                MessageBox.Show("Please select a pattern to edit.", "Selection Required",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void RemovePatternButton_Click(object sender, RoutedEventArgs e)
        {
            if (PatternsListBox.SelectedIndex >= 0)
            {
                _textFillPatterns.RemoveAt(PatternsListBox.SelectedIndex);
                RefreshPatternsList();
            }
            else
            {
                MessageBox.Show("Please select a pattern to remove.", "Selection Required",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ApplyFillButton_Click(object sender, RoutedEventArgs e)
        {
            if (_textFillPatterns.Count == 0)
            {
                MessageBox.Show("Please define at least one text pattern before applying.", "No Patterns Defined",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _callback?.Invoke(_textFillPatterns);
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    public class TextFillPattern
    {
        public string ContainsText { get; set; }
        public int SelectedOption { get; set; } // 1, 2, 3, or 4
    }
}