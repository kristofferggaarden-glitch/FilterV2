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

            RefreshGroupList();
        }

        private void RefreshGroupList()
        {
            GroupsListBox.Items.Clear();
            foreach (var group in _groupDefinitions.OrderBy(g => g.Priority))
            {
                GroupsListBox.Items.Add($"Priority {group.Priority}: {group.GroupName} (Contains: '{group.ContainsText}')");
            }
        }

        private void AddGroupButton_Click(object sender, RoutedEventArgs e)
        {
            string groupName = GroupNameTextBox.Text.Trim();
            string containsText = ContainsTextTextBox.Text.Trim();

            if (string.IsNullOrEmpty(groupName) || string.IsNullOrEmpty(containsText))
            {
                MessageBox.Show("Please enter both Group Name and Contains Text.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int priority = _groupDefinitions.Count > 0 ? _groupDefinitions.Max(g => g.Priority) + 1 : 1;

            _groupDefinitions.Add(new GroupDefinition
            {
                GroupName = groupName,
                ContainsText = containsText,
                Priority = priority
            });

            GroupNameTextBox.Clear();
            ContainsTextTextBox.Clear();
            RefreshGroupList();
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
                var orderedGroups = _groupDefinitions.OrderBy(g => g.Priority).ToList();
                var selectedGroup = orderedGroups[GroupsListBox.SelectedIndex];
                var previousGroup = orderedGroups[GroupsListBox.SelectedIndex - 1];

                int tempPriority = selectedGroup.Priority;
                selectedGroup.Priority = previousGroup.Priority;
                previousGroup.Priority = tempPriority;

                RefreshGroupList();
                GroupsListBox.SelectedIndex = GroupsListBox.SelectedIndex - 1;
            }
        }

        private void MoveDownButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsListBox.SelectedIndex >= 0 && GroupsListBox.SelectedIndex < GroupsListBox.Items.Count - 1)
            {
                var orderedGroups = _groupDefinitions.OrderBy(g => g.Priority).ToList();
                var selectedGroup = orderedGroups[GroupsListBox.SelectedIndex];
                var nextGroup = orderedGroups[GroupsListBox.SelectedIndex + 1];

                int tempPriority = selectedGroup.Priority;
                selectedGroup.Priority = nextGroup.Priority;
                nextGroup.Priority = tempPriority;

                RefreshGroupList();
                GroupsListBox.SelectedIndex = GroupsListBox.SelectedIndex + 1;
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