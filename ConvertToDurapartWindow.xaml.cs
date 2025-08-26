using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FilterV1
{
    public partial class ConvertToDurapartWindow : Window
    {
        private Action<List<ConversionRule>> _callback;
        private List<ConversionRule> _conversionRules;

        public ConvertToDurapartWindow(List<ConversionRule> existingRules, Action<List<ConversionRule>> callback)
        {
            InitializeComponent();
            _callback = callback;

            // Create a deep copy of existing rules
            _conversionRules = existingRules?.Select(r => new ConversionRule
            {
                FromText = r?.FromText ?? "",
                ToText = r?.ToText ?? ""
            }).ToList() ?? new List<ConversionRule>();

            RefreshRulesList();
        }

        private void RefreshRulesList()
        {
            ConversionRulesGrid.ItemsSource = null;
            ConversionRulesGrid.ItemsSource = _conversionRules;
        }

        private void FromTextTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ToTextTextBox.Focus(); // Move to next field
            }
        }

        private void ToTextTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddRuleButton_Click(sender, e);
            }
        }

        private void AddRuleButton_Click(object sender, RoutedEventArgs e)
        {
            string fromText = FromTextTextBox.Text.Trim();
            string toText = ToTextTextBox.Text.Trim();

            if (string.IsNullOrEmpty(fromText) || string.IsNullOrEmpty(toText))
            {
                MessageBox.Show("Please enter both 'From' and 'To' text.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Check for duplicates
            if (_conversionRules.Any(r => r.FromText.Equals(fromText, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("A rule with this 'From' text already exists.", "Duplicate Entry",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _conversionRules.Add(new ConversionRule
            {
                FromText = fromText,
                ToText = toText
            });

            FromTextTextBox.Clear();
            ToTextTextBox.Clear();
            FromTextTextBox.Focus(); // Keep focus for easy multiple entries
            RefreshRulesList();
        }

        private void EditRuleButton_Click(object sender, RoutedEventArgs e)
        {
            if (ConversionRulesGrid.SelectedIndex >= 0)
            {
                var selectedRule = _conversionRules[ConversionRulesGrid.SelectedIndex];

                // Load the selected rule into the input fields
                FromTextTextBox.Text = selectedRule.FromText;
                ToTextTextBox.Text = selectedRule.ToText;

                // Remove the old rule (it will be re-added when user clicks Add)
                _conversionRules.Remove(selectedRule);
                RefreshRulesList();

                FromTextTextBox.Focus();
            }
            else
            {
                MessageBox.Show("Please select a rule to edit.", "Selection Required",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void RemoveRuleButton_Click(object sender, RoutedEventArgs e)
        {
            if (ConversionRulesGrid.SelectedIndex >= 0)
            {
                var selectedRule = _conversionRules[ConversionRulesGrid.SelectedIndex];
                _conversionRules.Remove(selectedRule);
                RefreshRulesList();
            }
            else
            {
                MessageBox.Show("Please select a rule to remove.", "Selection Required",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void AddDefaultsButton_Click(object sender, RoutedEventArgs e)
        {
            var defaults = new List<ConversionRule>
            {
                new ConversionRule { FromText = "UNIBK1.0", ToText = "RDXBK1.0" },
                new ConversionRule { FromText = "UNIBK1.5", ToText = "RDXBK1.5" },
                new ConversionRule { FromText = "UNIBK2.5", ToText = "RDXBK2.5" },
                new ConversionRule { FromText = "UNIBK4.0", ToText = "RDXBK4.0" },
                new ConversionRule { FromText = "HYLSE 1.0", ToText = "RDXHYLSE 1.0" },
                new ConversionRule { FromText = "HYLSE 1.5", ToText = "RDXHYLSE 1.5" },
                new ConversionRule { FromText = "HYLSE 2.5", ToText = "RDXHYLSE 2.5" },
                new ConversionRule { FromText = "HYLSE 4.0", ToText = "RDXHYLSE 4.0" }
            };

            int addedCount = 0;
            foreach (var defaultRule in defaults)
            {
                if (!_conversionRules.Any(r => r.FromText.Equals(defaultRule.FromText, StringComparison.OrdinalIgnoreCase)))
                {
                    _conversionRules.Add(defaultRule);
                    addedCount++;
                }
            }

            RefreshRulesList();

            if (addedCount > 0)
            {
                MessageBox.Show($"Added {addedCount} default conversion rules.", "Defaults Added",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("All default rules already exist.", "No Changes",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ApplyConversionButton_Click(object sender, RoutedEventArgs e)
        {
            if (_conversionRules.Count == 0)
            {
                MessageBox.Show("Please define at least one conversion rule before applying.", "No Rules Defined",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _callback?.Invoke(_conversionRules);
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    public class ConversionRule
    {
        public string FromText { get; set; } = "";
        public string ToText { get; set; } = "";
    }
}