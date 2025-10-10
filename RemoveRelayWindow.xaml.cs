using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FilterV1
{
    public partial class RemoveRelayWindow : Window
    {
        private Action<List<RemoveRelayPattern>, List<RemoveRelayPattern>> _callback; // First param: all patterns for saving, Second: enabled patterns for processing
        private List<RemoveRelayPatternViewModel> _removeRelayPatterns;

        public RemoveRelayWindow(List<RemoveRelayPattern> existingPatterns, Action<List<RemoveRelayPattern>, List<RemoveRelayPattern>> callback)
        {
            InitializeComponent();
            _callback = callback;

            // Create a deep copy of existing patterns with ViewModel wrappers
            _removeRelayPatterns = existingPatterns?.Select(p => new RemoveRelayPatternViewModel
            {
                ContainsText = p?.ContainsText ?? "",
                // Default to disabled so that the user must explicitly enable patterns each session.
                // This corresponds to the requirement that irregular removal patterns should not
                // remain checked between sessions.
                IsEnabled = false
            }).ToList() ?? new List<RemoveRelayPatternViewModel>();

            RefreshPatternsList();
        }

        private void RefreshPatternsList()
        {
            PatternsListBox.ItemsSource = null;
            PatternsListBox.ItemsSource = _removeRelayPatterns;
        }

        private void ContainsTextTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddPatternButton_Click(sender, e);
            }
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

            // Handle multi-line paste - split by newlines and process each line
            string[] lines = containsText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int addedCount = 0;

            foreach (string line in lines)
            {
                string trimmedLine = line.Trim();
                if (string.IsNullOrEmpty(trimmedLine)) continue;

                // If line contains tabs (Excel multi-column paste), take only the first column
                if (trimmedLine.Contains('\t'))
                {
                    string[] parts = trimmedLine.Split('\t', StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length > 0)
                    {
                        trimmedLine = parts[0].Trim();
                    }
                }

                if (string.IsNullOrEmpty(trimmedLine)) continue;

                // Check for duplicates
                if (_removeRelayPatterns.Any(p => p.ContainsText.Equals(trimmedLine, StringComparison.OrdinalIgnoreCase)))
                {
                    continue; // Skip duplicates
                }

                _removeRelayPatterns.Add(new RemoveRelayPatternViewModel
                {
                    ContainsText = trimmedLine,
                    // New patterns are initially disabled.  Users must opt in by selecting the
                    // checkbox, aligning with the behaviour that patterns should not be applied
                    // automatically after being added.
                    IsEnabled = false
                });

                addedCount++;
            }

            ContainsTextTextBox.Clear();
            ContainsTextTextBox.Focus(); // Keep focus for easy multiple entries
            RefreshPatternsList();

            if (addedCount > 1)
            {
                MessageBox.Show($"Added {addedCount} patterns successfully.", "Multiple Patterns Added",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else if (addedCount == 0)
            {
                MessageBox.Show("No new patterns were added. Check for duplicates or empty text.", "No Patterns Added",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var pattern in _removeRelayPatterns)
            {
                pattern.IsEnabled = true;
            }
            PatternsListBox.Items.Refresh();
        }

        private void DeselectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var pattern in _removeRelayPatterns)
            {
                pattern.IsEnabled = false;
            }
            PatternsListBox.Items.Refresh();
        }

        private void RemovePatternButton_Click(object sender, RoutedEventArgs e)
        {
            if (PatternsListBox.SelectedIndex >= 0)
            {
                var selectedPattern = _removeRelayPatterns[PatternsListBox.SelectedIndex];
                _removeRelayPatterns.Remove(selectedPattern);
                RefreshPatternsList();
            }
            else
            {
                MessageBox.Show("Please select a pattern to remove.", "Selection Required",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ApplyRemoveButton_Click(object sender, RoutedEventArgs e)
        {
            var enabledPatterns = _removeRelayPatterns.Where(p => p.IsEnabled).ToList();

            if (enabledPatterns.Count == 0)
            {
                MessageBox.Show("Please enable at least one text pattern before applying.", "No Patterns Enabled",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Convert all patterns and enabled patterns to regular model for callback
            var allPatterns = _removeRelayPatterns.Select(p => new RemoveRelayPattern { ContainsText = p.ContainsText }).ToList();
            var enabledPatternsOnly = enabledPatterns.Select(p => new RemoveRelayPattern { ContainsText = p.ContainsText }).ToList();

            _callback?.Invoke(allPatterns, enabledPatternsOnly);
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    public class RemoveRelayPattern
    {
        public string ContainsText { get; set; } = "";
    }

    public class RemoveRelayPatternViewModel
    {
        public string ContainsText { get; set; } = "";
        public bool IsEnabled { get; set; } = true;
        public string DisplayText => $"Remove cells containing: '{ContainsText}'";
    }
}