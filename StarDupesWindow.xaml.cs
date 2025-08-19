using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FilterV1
{
    public partial class StarDupesWindow : Window
    {
        private Action<List<StarDupesRule>> _callback;
        private List<StarDupesRule> _starDupesRules;

        public StarDupesWindow(List<StarDupesRule> existingRules, Action<List<StarDupesRule>> callback)
        {
            InitializeComponent();
            _callback = callback;

            // Create a deep copy of existing rules with null safety
            _starDupesRules = existingRules?.Select(r => new StarDupesRule
            {
                DuplicateContains = r?.DuplicateContains ?? "",
                AdjacentContains = r?.AdjacentContains ?? "",
                Priority = r?.Priority ?? 0
            }).ToList() ?? new List<StarDupesRule>();

            // Ensure priorities are set correctly
            EnsurePrioritiesAreSet();

            RefreshRulesList();
        }

        private void EnsurePrioritiesAreSet()
        {
            // Check if any rules have priority 0 (unset) and assign them proper priorities
            var rulesWithoutPriority = _starDupesRules.Where(r => r.Priority == 0).ToList();

            if (rulesWithoutPriority.Any())
            {
                int maxPriority = _starDupesRules.Where(r => r.Priority > 0).Any() ?
                    _starDupesRules.Where(r => r.Priority > 0).Max(r => r.Priority) : 0;

                foreach (var rule in rulesWithoutPriority)
                {
                    rule.Priority = ++maxPriority;
                }
            }
        }

        private void RefreshRulesList()
        {
            RulesListBox.Items.Clear();
            foreach (var rule in _starDupesRules.Where(r => r != null).OrderBy(r => r.Priority))
            {
                string duplicateText = rule.DuplicateContains ?? "";
                string adjacentText = rule.AdjacentContains ?? "";
                RulesListBox.Items.Add($"Priority {rule.Priority}: Duplicate contains '{duplicateText}' + Adjacent contains '{adjacentText}'");
            }
        }

        private void DuplicateContainsTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AdjacentContainsTextBox.Focus(); // Move to next field
            }
        }

        private void AdjacentContainsTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddRuleButton_Click(sender, e);
            }
        }

        private void AddRuleButton_Click(object sender, RoutedEventArgs e)
        {
            string duplicateContains = DuplicateContainsTextBox.Text.Trim();
            string adjacentContains = AdjacentContainsTextBox.Text.Trim();

            if (string.IsNullOrEmpty(duplicateContains) || string.IsNullOrEmpty(adjacentContains))
            {
                MessageBox.Show("Please enter both 'Duplicate Contains' and 'Adjacent Contains' text.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Check for duplicates with null safety
            if (_starDupesRules.Any(r =>
                !string.IsNullOrEmpty(r.DuplicateContains) &&
                !string.IsNullOrEmpty(r.AdjacentContains) &&
                r.DuplicateContains.Equals(duplicateContains, StringComparison.OrdinalIgnoreCase) &&
                r.AdjacentContains.Equals(adjacentContains, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("A rule with this combination already exists.", "Duplicate Entry",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int priority = _starDupesRules.Count > 0 ? _starDupesRules.Max(r => r.Priority) + 1 : 1;

            _starDupesRules.Add(new StarDupesRule
            {
                DuplicateContains = duplicateContains,
                AdjacentContains = adjacentContains,
                Priority = priority
            });

            DuplicateContainsTextBox.Clear();
            AdjacentContainsTextBox.Clear();
            DuplicateContainsTextBox.Focus(); // Keep focus for easy multiple entries
            RefreshRulesList();
        }

        private void MoveUpButton_Click(object sender, RoutedEventArgs e)
        {
            if (RulesListBox.SelectedIndex > 0)
            {
                int selectedIndex = RulesListBox.SelectedIndex;
                var orderedRules = _starDupesRules.OrderBy(r => r.Priority).ToList();
                var selectedRule = orderedRules[selectedIndex];
                var previousRule = orderedRules[selectedIndex - 1];

                int tempPriority = selectedRule.Priority;
                selectedRule.Priority = previousRule.Priority;
                previousRule.Priority = tempPriority;

                RefreshRulesList();
                RulesListBox.SelectedIndex = selectedIndex - 1;
            }
        }

        private void MoveDownButton_Click(object sender, RoutedEventArgs e)
        {
            if (RulesListBox.SelectedIndex >= 0 && RulesListBox.SelectedIndex < RulesListBox.Items.Count - 1)
            {
                int selectedIndex = RulesListBox.SelectedIndex;
                var orderedRules = _starDupesRules.OrderBy(r => r.Priority).ToList();
                var selectedRule = orderedRules[selectedIndex];
                var nextRule = orderedRules[selectedIndex + 1];

                int tempPriority = selectedRule.Priority;
                selectedRule.Priority = nextRule.Priority;
                nextRule.Priority = tempPriority;

                RefreshRulesList();
                RulesListBox.SelectedIndex = selectedIndex + 1;
            }
        }

        private void EditRuleButton_Click(object sender, RoutedEventArgs e)
        {
            if (RulesListBox.SelectedIndex >= 0)
            {
                var orderedRules = _starDupesRules.OrderBy(r => r.Priority).ToList();
                var selectedRule = orderedRules[RulesListBox.SelectedIndex];

                // Load the selected rule into the input fields
                DuplicateContainsTextBox.Text = selectedRule.DuplicateContains;
                AdjacentContainsTextBox.Text = selectedRule.AdjacentContains;

                // Remove the old rule (it will be re-added when user clicks Add)
                _starDupesRules.Remove(selectedRule);
                RefreshRulesList();

                DuplicateContainsTextBox.Focus();
            }
            else
            {
                MessageBox.Show("Please select a rule to edit.", "Selection Required",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void RemoveRuleButton_Click(object sender, RoutedEventArgs e)
        {
            if (RulesListBox.SelectedIndex >= 0)
            {
                var orderedRules = _starDupesRules.OrderBy(r => r.Priority).ToList();
                var selectedRule = orderedRules[RulesListBox.SelectedIndex];
                _starDupesRules.Remove(selectedRule);
                RefreshRulesList();
            }
            else
            {
                MessageBox.Show("Please select a rule to remove.", "Selection Required",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ApplyStarDupesButton_Click(object sender, RoutedEventArgs e)
        {
            _callback?.Invoke(_starDupesRules);
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    public class StarDupesRule
    {
        public string DuplicateContains { get; set; } = "";
        public string AdjacentContains { get; set; } = "";
        public int Priority { get; set; } = 0;
    }
}