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
        // Stores the current search term entered by the user.  When non‑empty the group list
        // is filtered to only include groups whose ContainsText matches this term.  Filtering
        // is case‑insensitive.  Updates to this value trigger a refresh of the list.
        private string _searchTerm = string.Empty;

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

        /// <summary>
        /// Updates the current search term whenever the search text box changes and refreshes
        /// the group list.  Filtering is case‑insensitive and matches the ContainsText field
        /// of each group.  Clearing the search box will restore the full list.
        /// </summary>
        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _searchTerm = (sender as TextBox)?.Text ?? string.Empty;
            RefreshGroupList();
        }

        private void RefreshGroupList()
        {
            GroupsListBox.Items.Clear();
            var ordered = _groupDefinitions.OrderBy(g => g.Priority);
            foreach (var group in ordered)
            {
                // Apply search filter if present.  If the search term is blank we include all
                // groups.  Otherwise include only those whose ContainsText includes the term.
                if (!string.IsNullOrWhiteSpace(_searchTerm))
                {
                    if (group.ContainsText == null ||
                        group.ContainsText.IndexOf(_searchTerm, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }
                }
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
                // Find highest existing priority
                int maxPriority = _groupDefinitions.Any() ? _groupDefinitions.Max(g => g.Priority) : 0;

                // Assign priorities to the new entries starting from maxPriority + 1 (LOWEST priority)
                for (int i = 0; i < newEntries.Count; i++)
                {
                    newEntries[i].Priority = maxPriority + i + 1;
                }

                // Add the new entries to the collection
                _groupDefinitions.AddRange(newEntries);
            }

            // Clear the input and refresh the UI
            ContainsTextTextBox.Clear();
            ContainsTextTextBox.Focus(); // Keep focus for easy multiple entries
            RefreshGroupList();

            // After refreshing, scroll to the bottom of the list so the user can see the newly added group(s)
            if (GroupsListBox.Items.Count > 0)
            {
                var lastItem = GroupsListBox.Items[GroupsListBox.Items.Count - 1];
                GroupsListBox.ScrollIntoView(lastItem);
            }

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
            if (GroupsListBox.SelectedItem != null)
            {
                // Parse the ContainsText value from the selected list item.  Items are formatted as
                // "Priority X: Contains 'text'".  Extract the substring within single quotes.
                string item = GroupsListBox.SelectedItem.ToString();
                string contains = null;
                int start = item.IndexOf("Contains '");
                int end = item.LastIndexOf("'");
                if (start >= 0 && end > start + 10)
                {
                    contains = item.Substring(start + 10, end - (start + 10));
                }
                if (!string.IsNullOrWhiteSpace(contains))
                {
                    var group = _groupDefinitions.FirstOrDefault(g => string.Equals(g.ContainsText, contains, StringComparison.OrdinalIgnoreCase));
                    if (group != null)
                    {
                        _groupDefinitions.Remove(group);
                        RefreshGroupList();
                        return;
                    }
                }
                // Fallback: remove by index if parsing failed
                int index = GroupsListBox.SelectedIndex;
                if (index >= 0)
                {
                    var ordered = _groupDefinitions.OrderBy(g => g.Priority).ToList();
                    if (index < ordered.Count)
                    {
                        _groupDefinitions.Remove(ordered[index]);
                        RefreshGroupList();
                        return;
                    }
                }
            }
            MessageBox.Show("Please select a group to remove.", "Selection Required",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void MoveUpButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsListBox.SelectedItem != null)
            {
                // Identify the underlying group from the selected item using parsing
                string item = GroupsListBox.SelectedItem.ToString();
                string contains = null;
                int start = item.IndexOf("Contains '");
                int end = item.LastIndexOf("'");
                if (start >= 0 && end > start + 10)
                {
                    contains = item.Substring(start + 10, end - (start + 10));
                }
                GroupDefinition selectedGroup = null;
                if (!string.IsNullOrWhiteSpace(contains))
                {
                    selectedGroup = _groupDefinitions.FirstOrDefault(g => string.Equals(g.ContainsText, contains, StringComparison.OrdinalIgnoreCase));
                }
                // Fallback: derive by index in filtered list if parsing fails
                if (selectedGroup == null)
                {
                    int selIndex = GroupsListBox.SelectedIndex;
                    var visible = _groupDefinitions.OrderBy(g => g.Priority)
                        .Where(g => string.IsNullOrWhiteSpace(_searchTerm) ||
                                    (g.ContainsText != null && g.ContainsText.IndexOf(_searchTerm, StringComparison.OrdinalIgnoreCase) >= 0))
                        .ToList();
                    if (selIndex >= 0 && selIndex < visible.Count)
                    {
                        selectedGroup = visible[selIndex];
                    }
                }
                if (selectedGroup != null)
                {
                    // Find previous group in full ordered list (by priority)
                    var ordered = _groupDefinitions.OrderBy(g => g.Priority).ToList();
                    int currentIndex = ordered.IndexOf(selectedGroup);
                    if (currentIndex > 0)
                    {
                        var previousGroup = ordered[currentIndex - 1];
                        int tempPriority = selectedGroup.Priority;
                        selectedGroup.Priority = previousGroup.Priority;
                        previousGroup.Priority = tempPriority;
                        RefreshGroupList();
                        // Restore selection on the moved group in filtered view
                        SelectGroupInList(selectedGroup);
                    }
                }
            }
        }

        private void MoveDownButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsListBox.SelectedItem != null)
            {
                // Identify the underlying group from the selected item
                string item = GroupsListBox.SelectedItem.ToString();
                string contains = null;
                int start = item.IndexOf("Contains '");
                int end = item.LastIndexOf("'");
                if (start >= 0 && end > start + 10)
                {
                    contains = item.Substring(start + 10, end - (start + 10));
                }
                GroupDefinition selectedGroup = null;
                if (!string.IsNullOrWhiteSpace(contains))
                {
                    selectedGroup = _groupDefinitions.FirstOrDefault(g => string.Equals(g.ContainsText, contains, StringComparison.OrdinalIgnoreCase));
                }
                if (selectedGroup == null)
                {
                    int selIndex = GroupsListBox.SelectedIndex;
                    var visible = _groupDefinitions.OrderBy(g => g.Priority)
                        .Where(g => string.IsNullOrWhiteSpace(_searchTerm) ||
                                    (g.ContainsText != null && g.ContainsText.IndexOf(_searchTerm, StringComparison.OrdinalIgnoreCase) >= 0))
                        .ToList();
                    if (selIndex >= 0 && selIndex < visible.Count)
                    {
                        selectedGroup = visible[selIndex];
                    }
                }
                if (selectedGroup != null)
                {
                    var ordered = _groupDefinitions.OrderBy(g => g.Priority).ToList();
                    int currentIndex = ordered.IndexOf(selectedGroup);
                    if (currentIndex >= 0 && currentIndex < ordered.Count - 1)
                    {
                        var nextGroup = ordered[currentIndex + 1];
                        int tempPriority = selectedGroup.Priority;
                        selectedGroup.Priority = nextGroup.Priority;
                        nextGroup.Priority = tempPriority;
                        RefreshGroupList();
                        SelectGroupInList(selectedGroup);
                    }
                }
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

        /// <summary>
        /// Helper method to select a given group in the list after the list has been refreshed.  It
        /// searches for the list item string corresponding to the group's ContainsText value and
        /// selects it.  If the group does not appear in the filtered view (e.g. due to search), the
        /// selection is cleared.
        /// </summary>
        /// <param name="group">The group to select.</param>
        private void SelectGroupInList(GroupDefinition group)
        {
            if (group == null)
                return;
            string target = $"{group.ContainsText}";
            for (int i = 0; i < GroupsListBox.Items.Count; i++)
            {
                string item = GroupsListBox.Items[i]?.ToString();
                if (item != null && item.IndexOf($"'{target}'", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    GroupsListBox.SelectedIndex = i;
                    GroupsListBox.ScrollIntoView(GroupsListBox.Items[i]);
                    return;
                }
            }
            // Clear selection if not found
            GroupsListBox.SelectedIndex = -1;
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

            // Build a list of selected groups by parsing the selected list items.  Items are
            // formatted as "Priority X: Contains 'text'".  Extract 'text' and match against
            // the underlying _groupDefinitions list (case‑insensitive).
            var selectedGroups = new List<GroupDefinition>();
            foreach (var item in GroupsListBox.SelectedItems)
            {
                string str = item?.ToString();
                if (string.IsNullOrEmpty(str)) continue;
                string contains = null;
                int startIdx = str.IndexOf("Contains '");
                int endIdx = str.LastIndexOf("'");
                if (startIdx >= 0 && endIdx > startIdx + 10)
                {
                    contains = str.Substring(startIdx + 10, endIdx - (startIdx + 10));
                }
                if (!string.IsNullOrWhiteSpace(contains))
                {
                    var grp = _groupDefinitions.FirstOrDefault(g => string.Equals(g.ContainsText, contains, StringComparison.OrdinalIgnoreCase));
                    if (grp != null && !selectedGroups.Contains(grp))
                    {
                        selectedGroups.Add(grp);
                    }
                }
            }
            if (selectedGroups.Count == 0)
            {
                return;
            }

            // Get the currently ordered groups
            var ordered = _groupDefinitions.OrderBy(g => g.Priority).ToList();
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
                SelectGroupInList(grp);
                // Since SelectGroupInList selects and scrolls to a single item, subsequent calls
                // will override previous selections.  To allow multi‑selection of moved items we
                // manually add the selection after the call.
                for (int i = 0; i < GroupsListBox.Items.Count; i++)
                {
                    string itm = GroupsListBox.Items[i]?.ToString();
                    if (itm != null && itm.IndexOf($"'{grp.ContainsText}'", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        if (!GroupsListBox.SelectedItems.Contains(GroupsListBox.Items[i]))
                        {
                            GroupsListBox.SelectedItems.Add(GroupsListBox.Items[i]);
                        }
                        break;
                    }
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