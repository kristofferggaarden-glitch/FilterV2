using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace FilterV1
{
    public partial class CustomGroupWindow : Window
    {
        private Action<List<GroupDefinition>> _callback;
        private List<GroupDefinition> _groupDefinitions;

        public CustomGroupWindow(List<GroupDefinition> existingGroups, Action<List<GroupDefinition>> callback)
        {
            InitializeComponent();
            _callback = callback;

            // Create a deep copy of existing groups to avoid modifying the original until Apply is clicked
            _groupDefinitions = existingGroups.Select(g => new GroupDefinition
            {
                GroupName = g.GroupName,
                ContainsText = g.ContainsText,
                Priority = g.Priority
            }).ToList();

            // Enable multi-selection for the ListBox
            GroupsListBox.SelectionMode = SelectionMode.Extended;

            // Add keyboard support for Delete key
            GroupsListBox.KeyDown += GroupsListBox_KeyDown;

            RefreshGroupList();
        }

        private void GroupsListBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Delete)
            {
                RemoveGroupButton_Click(sender, new RoutedEventArgs());
            }
        }

        private void RefreshGroupList()
        {
            GroupsListBox.Items.Clear();
            foreach (var group in _groupDefinitions.OrderBy(g => g.Priority))
            {
                GroupsListBox.Items.Add($"Priority {group.Priority}: Contains '{group.ContainsText}'");
            }
        }

        private void ContainsTextTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                AddGroupButton_Click(sender, e);
            }
        }

        private void AddGroupButton_Click(object sender, RoutedEventArgs e)
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
                if (_groupDefinitions.Any(g => g.ContainsText.Equals(trimmedLine, StringComparison.OrdinalIgnoreCase)))
                {
                    continue; // Skip duplicates
                }

                int priority = _groupDefinitions.Count > 0 ? _groupDefinitions.Max(g => g.Priority) + 1 : 1;

                // Auto-generate group name based on contains text
                string groupName = $"Group {trimmedLine}";

                _groupDefinitions.Add(new GroupDefinition
                {
                    GroupName = groupName,
                    ContainsText = trimmedLine,
                    Priority = priority
                });

                addedCount++;
            }

            ContainsTextTextBox.Clear();
            ContainsTextTextBox.Focus(); // Keep focus for easy multiple entries
            RefreshGroupList();

            if (addedCount > 1)
            {
                MessageBox.Show($"Added {addedCount} groups successfully.", "Multiple Groups Added",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else if (addedCount == 0)
            {
                MessageBox.Show("No new groups were added. Check for duplicates or empty text.", "No Groups Added",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void RemoveGroupButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsListBox.SelectedIndex >= 0)
            {
                var selectedGroup = _groupDefinitions.OrderBy(g => g.Priority).ToList()[GroupsListBox.SelectedIndex];
                _groupDefinitions.Remove(selectedGroup);
                RefreshGroupList();
            }
            else
            {
                MessageBox.Show("Please select a group to remove.", "Selection Required",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void MoveUpButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsListBox.SelectedIndex > 0)
            {
                int selectedIndex = GroupsListBox.SelectedIndex;
                var orderedGroups = _groupDefinitions.OrderBy(g => g.Priority).ToList();
                var selectedGroup = orderedGroups[selectedIndex];
                var previousGroup = orderedGroups[selectedIndex - 1];

                int tempPriority = selectedGroup.Priority;
                selectedGroup.Priority = previousGroup.Priority;
                previousGroup.Priority = tempPriority;

                RefreshGroupList();
                GroupsListBox.SelectedIndex = selectedIndex - 1;
            }
        }

        private void MoveDownButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsListBox.SelectedIndex >= 0 && GroupsListBox.SelectedIndex < GroupsListBox.Items.Count - 1)
            {
                int selectedIndex = GroupsListBox.SelectedIndex;
                var orderedGroups = _groupDefinitions.OrderBy(g => g.Priority).ToList();
                var selectedGroup = orderedGroups[selectedIndex];
                var nextGroup = orderedGroups[selectedIndex + 1];

                int tempPriority = selectedGroup.Priority;
                selectedGroup.Priority = nextGroup.Priority;
                nextGroup.Priority = tempPriority;

                RefreshGroupList();
                GroupsListBox.SelectedIndex = selectedIndex + 1;
            }
        }

        private void ApplyGroupsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_groupDefinitions.Count == 0)
            {
                MessageBox.Show("Please define at least one group before applying.", "No Groups Defined",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _callback?.Invoke(_groupDefinitions);
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    public class GroupDefinition
    {
        public string GroupName { get; set; }
        public string ContainsText { get; set; }
        public int Priority { get; set; }
    }
}