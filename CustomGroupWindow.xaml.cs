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

            // Enable multi-selection for the ListBox.  Users can select multiple rows
            // and move them together using the new priority control.
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
            var newEntries = new List<GroupDefinition>();

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
                if (_groupDefinitions.Any(g => g.ContainsText.Equals(trimmedLine, StringComparison.OrdinalIgnoreCase)) ||
                    newEntries.Any(g => g.ContainsText.Equals(trimmedLine, StringComparison.OrdinalIgnoreCase)))
                {
                    continue; // Skip duplicates
                }

                // Auto-generate group name based on contains text
                string groupName = $"Group {trimmedLine}";

                newEntries.Add(new GroupDefinition
                {
                    GroupName = groupName,
                    ContainsText = trimmedLine,
                    Priority = 0 // Placeholder, will assign below
                });
            }

            if (newEntries.Count > 0)
            {
                // Shift existing group priorities down to make room for new entries at the top
                foreach (var g in _groupDefinitions)
                {
                    g.Priority += newEntries.Count;
                }

                // Assign priorities to the new entries starting from 1
                for (int i = 0; i < newEntries.Count; i++)
                {
                    newEntries[i].Priority = i + 1;
                }

                // Add the new entries to the collection
                _groupDefinitions.AddRange(newEntries);
            }

            // Clear the input and refresh the UI
            ContainsTextTextBox.Clear();
            ContainsTextTextBox.Focus(); // Keep focus for easy multiple entries
            RefreshGroupList();

            if (newEntries.Count > 1)
            {
                MessageBox.Show($"Added {newEntries.Count} groups successfully.", "Multiple Groups Added",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else if (newEntries.Count == 0)
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

        /// <summary>
        /// Handles the click for the Set Priority button. Moves all selected groups
        /// to the specified priority position. If the priority is invalid or no
        /// items are selected, an informational dialog is shown. After moving,
        /// group priorities are normalized (1..N) and the list is refreshed.
        /// </summary>
        private void SetPriorityButton_Click(object sender, RoutedEventArgs e)
        {
            // Ensure there is at least one selection
            if (GroupsListBox.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one group to move.", "Selection Required",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            // Parse desired priority from text box
            if (!int.TryParse(MoveToPriorityTextBox.Text.Trim(), out int targetPriority) || targetPriority < 1)
            {
                MessageBox.Show("Please enter a valid target priority (integer >= 1).", "Invalid Priority",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Get the currently ordered groups
            var ordered = _groupDefinitions.OrderBy(g => g.Priority).ToList();
            // Map selected items to group definitions. The ListBox displays strings; to map to
            // underlying objects we use the same ordered list by index.
            var selectedGroups = new List<GroupDefinition>();
            foreach (var item in GroupsListBox.SelectedItems)
            {
                int index = GroupsListBox.Items.IndexOf(item);
                if (index >= 0 && index < ordered.Count)
                {
                    selectedGroups.Add(ordered[index]);
                }
            }
            if (selectedGroups.Count == 0)
            {
                return;
            }

            // Remove selected groups from the ordered list
            ordered = ordered.Except(selectedGroups).ToList();
            // Clamp target priority within bounds (1..Count+1)
            if (targetPriority > ordered.Count + 1)
            {
                targetPriority = ordered.Count + 1;
            }
            // Insert selected groups at the desired position (1-based index converted to 0-based)
            ordered.InsertRange(targetPriority - 1, selectedGroups);

            // Reassign sequential priorities
            for (int i = 0; i < ordered.Count; i++)
            {
                ordered[i].Priority = i + 1;
            }

            // Replace the original collection with the reordered list
            _groupDefinitions = ordered;

            // Refresh UI and reselect moved items in their new positions
            RefreshGroupList();
            GroupsListBox.SelectedItems.Clear();
            foreach (var grp in selectedGroups)
            {
                int newIndex = ordered.IndexOf(grp);
                if (newIndex >= 0 && newIndex < GroupsListBox.Items.Count)
                {
                    // The ListBox items correspond one-to-one with ordered groups
                    GroupsListBox.SelectedItems.Add(GroupsListBox.Items[newIndex]);
                }
            }
        }
    }

    public class GroupDefinition
    {
        public string GroupName { get; set; }
        public string ContainsText { get; set; }
        public int Priority { get; set; }
    }
}