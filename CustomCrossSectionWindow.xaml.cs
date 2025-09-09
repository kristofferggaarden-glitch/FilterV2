using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FilterV1
{
    /// <summary>
    /// Interaction logic for CustomCrossSectionWindow.xaml
    /// This window allows users to define custom cross‑section assignments for pairs of values in columns 5 and 6.
    /// Users can paste multiple rows from Excel, set tverrsnitt options per row, and the selections will override standard tverrsnitt.
    /// </summary>
    public partial class CustomCrossSectionWindow : Window
    {
        public class CrossRow
        {
            public string Col5Text { get; set; } = string.Empty;
            public string Col6Text { get; set; } = string.Empty;
            public int SelectedOption { get; set; } = 1;
        }

        private readonly Action<List<CrossRow>> _callback;
        private readonly List<CrossRow> _rows;
        private readonly List<CrossRow> _originalRows;

        public CustomCrossSectionWindow(List<CrossRow> existingRows, Action<List<CrossRow>> callback)
        {
            InitializeComponent();
            _callback = callback;
            _rows = existingRows != null ? existingRows.Select(r => new CrossRow
            {
                Col5Text = r.Col5Text,
                Col6Text = r.Col6Text,
                SelectedOption = r.SelectedOption
            }).ToList() : new List<CrossRow>();
            _originalRows = existingRows != null ? existingRows.Select(r => new CrossRow
            {
                Col5Text = r.Col5Text,
                Col6Text = r.Col6Text,
                SelectedOption = r.SelectedOption
            }).ToList() : new List<CrossRow>();
            CrossRowsGrid.ItemsSource = _rows;
        }

        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (PasteTextBox.Text == "Lim inn celler her...")
            {
                PasteTextBox.Text = string.Empty;
            }
        }

        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(PasteTextBox.Text))
            {
                PasteTextBox.Text = "Lim inn celler her...";
            }
        }

        private void PasteTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddButton_Click(sender, e);
            }
        }

        private void PasteTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                string clipboardText = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(clipboardText))
                {
                    MessageBox.Show("Ingen tekst i utklippstavlen.", "Feil", MessageBoxButton.OK, MessageBoxImage.Warning);
                    e.Handled = true;
                    return;
                }
                bool added = false;
                foreach (var row in ParseLinesToRows(clipboardText))
                {
                    _rows.Add(row);
                    added = true;
                }
                if (added)
                {
                    CrossRowsGrid.Items.Refresh();
                    PasteTextBox.Text = string.Empty;
                    PasteTextBox.Focus();
                }
                e.Handled = true;
            }
        }

        private List<CrossRow> ParseLinesToRows(string text)
        {
            var results = new List<CrossRow>();
            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;
                string[] parts = trimmed.Contains('\t') ? trimmed.Split('\t') : trimmed.Split(' ');
                if (parts.Length >= 2)
                {
                    results.Add(new CrossRow
                    {
                        Col5Text = parts[0].Trim(),
                        Col6Text = parts[1].Trim(),
                        SelectedOption = 1
                    });
                }
            }
            return results;
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            string pastedText = PasteTextBox.Text.Trim();
            if (string.IsNullOrEmpty(pastedText) || pastedText == "Lim inn celler her...")
            {
                MessageBox.Show("Lim inn gyldig tekst.", "Feil", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            var newRows = ParseLinesToRows(pastedText);
            if (newRows.Count > 0)
            {
                foreach (var row in newRows)
                {
                    _rows.Add(row);
                }
                CrossRowsGrid.Items.Refresh();
                PasteTextBox.Text = string.Empty;
                PasteTextBox.Focus();
            }
            else
            {
                MessageBox.Show("Fant ingen gyldige rader.", "Feil", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void RemoveSelectedButton_Click(object sender, RoutedEventArgs e)
        {
            var selected = CrossRowsGrid.SelectedItems.Cast<CrossRow>().ToList();
            foreach (var row in selected)
            {
                _rows.Remove(row);
            }
            CrossRowsGrid.Items.Refresh();
        }

        private void ClearAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Er du sikker på at du vil fjerne alle radene?", "Bekreft", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                _rows.Clear();
                CrossRowsGrid.Items.Refresh();
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            // Use a copy to avoid shared references
            _callback?.Invoke(_rows.Select(r => new CrossRow
            {
                Col5Text = r.Col5Text,
                Col6Text = r.Col6Text,
                SelectedOption = r.SelectedOption
            }).ToList());
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            // Return original list (no changes)
            _callback?.Invoke(_originalRows.Select(r => new CrossRow
            {
                Col5Text = r.Col5Text,
                Col6Text = r.Col6Text,
                SelectedOption = r.SelectedOption
            }).ToList());
            this.Close();
        }
    }
}