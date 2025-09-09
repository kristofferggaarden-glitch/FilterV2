using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace FilterV1
{
    /// <summary>
    /// Interaction logic for RisingNumbersOptionsWindow.xaml
    /// Provides a UI for specifying exceptions when removing rising number pairs.
    /// </summary>
    public partial class RisingNumbersOptionsWindow : Window
    {
        private readonly Action<List<string>> _callback;
        private readonly List<string> _originalExceptions;

        public RisingNumbersOptionsWindow(List<string> existingExceptions, Action<List<string>> callback)
        {
            InitializeComponent();
            _callback = callback;
            _originalExceptions = existingExceptions != null ? new List<string>(existingExceptions) : new List<string>();
            PopulateListBox();
        }

        private void PopulateListBox()
        {
            ExceptionsListBox.Items.Clear();
            foreach (var ex in _originalExceptions)
            {
                var cb = new CheckBox { Content = ex, IsChecked = true, Tag = ex };
                ExceptionsListBox.Items.Add(cb);
            }
        }

        private void AddExceptionButton_Click(object sender, RoutedEventArgs e)
        {
            var txt = NewExceptionTextBox.Text.Trim();
            if (string.IsNullOrEmpty(txt)) return;
            // Avoid duplicates; just ensure existing item is checked
            var existing = _originalExceptions.FirstOrDefault(s => s.Equals(txt, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                foreach (CheckBox cb in ExceptionsListBox.Items)
                {
                    if (cb.Tag is string s && s.Equals(existing, StringComparison.OrdinalIgnoreCase))
                    {
                        cb.IsChecked = true;
                        break;
                    }
                }
            }
            else
            {
                _originalExceptions.Add(txt);
                var cb = new CheckBox { Content = txt, IsChecked = true, Tag = txt };
                ExceptionsListBox.Items.Add(cb);
            }
            NewExceptionTextBox.Clear();
            NewExceptionTextBox.Focus();
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (CheckBox cb in ExceptionsListBox.Items)
                cb.IsChecked = true;
        }

        private void DeselectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (CheckBox cb in ExceptionsListBox.Items)
                cb.IsChecked = false;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            var selected = new List<string>();
            foreach (CheckBox cb in ExceptionsListBox.Items)
            {
                if (cb.IsChecked == true && cb.Tag is string s)
                {
                    selected.Add(s);
                }
            }
            _callback?.Invoke(selected);
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            // Return original exceptions on cancel
            _callback?.Invoke(new List<string>(_originalExceptions));
            this.Close();
        }
    }
}