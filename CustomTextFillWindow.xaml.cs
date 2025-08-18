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

            // Create a deep copy of existing patterns and ensure they have priorities
            _textFillPatterns = existingPatterns.Select(p => new TextFillPattern
            {
                ContainsText = p.ContainsText,
                SelectedOption = p.SelectedOption,
                Priority = p.Priority
            }).ToList();

            // Fix any patterns that don't have priorities assigned
            EnsurePrioritiesAreSet();

            OptionComboBox.SelectionChanged += OptionComboBox_SelectionChanged;
            RefreshPatternsList();
            UpdatePreview();
        }

        private void EnsurePrioritiesAreSet()
        {
            // Check if any patterns have priority 0 (unset) and assign them proper priorities
            var patternsWithoutPriority = _textFillPatterns.Where(p => p.Priority == 0).ToList();

            if (patternsWithoutPriority.Any())
            {
                int maxPriority = _textFillPatterns.Where(p => p.Priority > 0).Any() ?
                    _textFillPatterns.Where(p => p.Priority > 0).Max(p => p.Priority) : 0;

                foreach (var pattern in patternsWithoutPriority)
                {
                    pattern.Priority = ++maxPriority;
                }
            }
        }

        private void RefreshPatternsList()
        {
            PatternsListBox.Items.Clear();
            foreach (var pattern in _textFillPatterns.OrderBy(p => p.Priority))
            {
                string optionText = GetOptionDescription(pattern.SelectedOption);
                PatternsListBox.Items.Add($"Priority {pattern.Priority}: '{pattern.ContainsText}' → {optionText}");
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

            int priority = _textFillPatterns.Count > 0 ? _textFillPatterns.Max(p => p.Priority) + 1 : 1;

            _textFillPatterns.Add(new TextFillPattern
            {
                ContainsText = containsText,
                SelectedOption = selectedOption,
                Priority = priority
            });

            ContainsTextTextBox.Clear();
            ContainsTextTextBox.Focus(); // Keep focus for easy multiple entries
            RefreshPatternsList();
        }

        private void MoveUpButton_Click(object sender, RoutedEventArgs e)
        {
            if (PatternsListBox.SelectedIndex > 0)
            {
                int selectedIndex = PatternsListBox.SelectedIndex;
                var orderedPatterns = _textFillPatterns.OrderBy(p => p.Priority).ToList();
                var selectedPattern = orderedPatterns[selectedIndex];
                var previousPattern = orderedPatterns[selectedIndex - 1];

                int tempPriority = selectedPattern.Priority;
                selectedPattern.Priority = previousPattern.Priority;
                previousPattern.Priority = tempPriority;

                RefreshPatternsList();
                PatternsListBox.SelectedIndex = selectedIndex - 1;
            }
        }

        private void MoveDownButton_Click(object sender, RoutedEventArgs e)
        {
            if (PatternsListBox.SelectedIndex >= 0 && PatternsListBox.SelectedIndex < PatternsListBox.Items.Count - 1)
            {
                int selectedIndex = PatternsListBox.SelectedIndex;
                var orderedPatterns = _textFillPatterns.OrderBy(p => p.Priority).ToList();
                var selectedPattern = orderedPatterns[selectedIndex];
                var nextPattern = orderedPatterns[selectedIndex + 1];

                int tempPriority = selectedPattern.Priority;
                selectedPattern.Priority = nextPattern.Priority;
                nextPattern.Priority = tempPriority;

                RefreshPatternsList();
                PatternsListBox.SelectedIndex = selectedIndex + 1;
            }
        }

        private void EditPatternButton_Click(object sender, RoutedEventArgs e)
        {
            if (PatternsListBox.SelectedIndex >= 0)
            {
                var orderedPatterns = _textFillPatterns.OrderBy(p => p.Priority).ToList();
                var selectedPattern = orderedPatterns[PatternsListBox.SelectedIndex];

                // Load the selected pattern into the input fields
                ContainsTextTextBox.Text = selectedPattern.ContainsText;
                OptionComboBox.SelectedIndex = selectedPattern.SelectedOption - 1; // Convert to 0-based index

                // Remove the old pattern (it will be re-added when user clicks Add)
                _textFillPatterns.Remove(selectedPattern);
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
                var orderedPatterns = _textFillPatterns.OrderBy(p => p.Priority).ToList();
                var selectedPattern = orderedPatterns[PatternsListBox.SelectedIndex];
                _textFillPatterns.Remove(selectedPattern);
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
        public int Priority { get; set; }
    }
}