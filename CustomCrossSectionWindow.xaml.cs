using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FilterV1
{
    /// <summary>
    /// Interaction logic for CustomCrossSectionWindow.xaml
    ///
    /// This window provides a user experience similar to the "Remove Cell Pairs" dialog.  Users can paste two
    /// columns of text (commonly column 5 and 6) from Excel or from the application's data preview, and
    /// assign custom cross‑sections (tverrsnitt) to those pairs.  Multiple rows may be selected and assigned
    /// via keyboard shortcuts (1..4) or context menu items.  The defined mappings are persisted and
    /// returned via a callback when the user clicks the "Apply Custom" button.
    /// </summary>
    public partial class CustomCrossSectionWindow : Window
    {
        /// <summary>
        /// Represents one user defined mapping consisting of two text values (from columns 5 and 6) and a
        /// selected option representing the desired cross‑section.  The SelectedOption property values map
        /// to the options list defined in the Tag property of the window (1=1.0, 2=1.5, 3=2.5, 4=4.0).
        /// </summary>
        public class CrossRow
        {
            public string Col5Text { get; set; } = string.Empty;
            public string Col6Text { get; set; } = string.Empty;
            public int SelectedOption { get; set; } = 1;
        }

        /// <summary>
        /// Represents an option item for binding to ComboBoxes.  Each option has a numeric value and a
        /// user friendly label.  The numeric value corresponds to the SelectedOption on CrossRow.
        /// </summary>
        public class OptionItem
        {
            public int Value { get; set; }
            public string Label { get; set; } = string.Empty;
        }

        private readonly Action<List<CrossRow>> _onApply;
        private readonly string _storeFile;
        private ObservableCollection<CrossRow> _rows = new ObservableCollection<CrossRow>();

        /// <summary>
        /// Constructs a new custom cross section window.  If an existing list of rows is provided, it will be
        /// used as the initial set of mappings; otherwise the list will be loaded from persistent storage.
        /// The callback <paramref name="onApply"/> will be invoked with the current list when the user applies
        /// the custom mappings.
        /// </summary>
        /// <param name="existing">An optional list of previously defined rows.</param>
        /// <param name="onApply">A callback invoked when the user applies the custom mappings.</param>
        public CustomCrossSectionWindow(List<CrossRow>? existing, Action<List<CrossRow>> onApply)
        {
            InitializeComponent();
            _onApply = onApply;

            // Build storage file path under %AppData%\FilterV1
            _storeFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FilterV1", "custom_cross_sections.json");

            // Populate the Tag property on the window with a list of available cross‑section options.
            // This list is used by the DataGridComboBoxColumn and ensures a single source of truth.
            this.Tag = new List<OptionItem>
            {
                new OptionItem { Value = 1, Label = "1.0 mm²" },
                new OptionItem { Value = 2, Label = "1.5 mm²" },
                new OptionItem { Value = 3, Label = "2.5 mm²" },
                new OptionItem { Value = 4, Label = "4.0 mm²" },
            };

            // Load rows either from the provided list or from persistent storage.
            Load(existing);
            CrossGrid.ItemsSource = _rows;

            // Handle keyboard shortcuts at Window level
            this.PreviewKeyDown += Window_PreviewKeyDown;
        }

        /// <summary>
        /// Loads rows from the provided list or from persistent storage.  If both sources are empty,
        /// the internal list is cleared.  Any extraneous whitespace is trimmed and invalid SelectedOption
        /// values are defaulted to 1.
        /// </summary>
        /// <param name="existing">Optional list of rows to load.</param>
        private void Load(List<CrossRow>? existing)
        {
            try
            {
                List<CrossRow> start;
                if (existing != null && existing.Count > 0)
                {
                    start = existing;
                }
                else if (File.Exists(_storeFile))
                {
                    var json = File.ReadAllText(_storeFile);
                    start = JsonSerializer.Deserialize<List<CrossRow>>(json) ?? new List<CrossRow>();
                }
                else
                {
                    start = new List<CrossRow>();
                }

                _rows.Clear();
                foreach (var r in start)
                {
                    _rows.Add(new CrossRow
                    {
                        Col5Text = r.Col5Text?.Trim() ?? string.Empty,
                        Col6Text = r.Col6Text?.Trim() ?? string.Empty,
                        SelectedOption = (r.SelectedOption >= 1 && r.SelectedOption <= 4) ? r.SelectedOption : 1
                    });
                }
            }
            catch
            {
                _rows.Clear();
            }
        }

        /// <summary>
        /// Persists the current list of rows to the storage file.  Errors are silently ignored.
        /// </summary>
        private void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_storeFile)!);
                File.WriteAllText(_storeFile, JsonSerializer.Serialize(_rows.ToList(), new JsonSerializerOptions { WriteIndented = true }));
            }
            catch
            {
                // Ignore any errors during save
            }
        }

        /// <summary>
        /// When the paste box gains focus, remove the placeholder text if present.
        /// </summary>
        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (PasteTextBox.Text == "Paste cell pairs here...")
            {
                PasteTextBox.Text = string.Empty;
            }
        }

        /// <summary>
        /// When the paste box loses focus and is empty, restore the placeholder text.
        /// </summary>
        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(PasteTextBox.Text))
            {
                PasteTextBox.Text = "Paste cell pairs here...";
            }
        }

        /// <summary>
        /// Handles the Enter key in the paste box.  When pressed, the current text is parsed and added
        /// as new rows, then the textbox is cleared.  The default event handling is suppressed.
        /// </summary>
        private void PasteTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                AddFromRawText(PasteTextBox.Text);
                PasteTextBox.Text = string.Empty;
            }
        }

        /// <summary>
        /// Intercepts Ctrl+V in the paste box to allow proper parsing of multi‑line clipboard data.
        /// When Ctrl+V is pressed, the clipboard text is parsed and added as new rows.  The paste
        /// operation into the TextBox is suppressed.
        /// </summary>
        private void PasteTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                string clipboardText = Clipboard.GetText();
                if (!string.IsNullOrWhiteSpace(clipboardText))
                {
                    AddFromRawText(clipboardText);
                    PasteTextBox.Text = string.Empty;
                }
                e.Handled = true;
            }
        }

        /// <summary>
        /// Splits a single line into parts, removing any quotes and splitting on TAB, semicolon, or comma.
        /// </summary>
        private static string[] SplitLine(string line)
        {
            var cleaned = (line ?? string.Empty).Replace("\"", string.Empty).Trim();
            return cleaned.Split(new[] { '\t', ';', ',' }, StringSplitOptions.None);
        }

        /// <summary>
        /// Parses arbitrary multi‑line text into pairs of values and adds them to the internal list.  Empty
        /// lines are ignored, and if both values on a line are empty the line is skipped.  Duplicate
        /// mappings (case‑insensitive comparison of both columns) are not added again.
        /// </summary>
        private void AddFromRawText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return;

            string[] lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int added = 0;

            // To ensure that the newest entries appear at the top while preserving
            // the order the user pasted them, process the lines in reverse order
            // and insert each new row at index 0. Without reversing, the final
            // order in the collection would be the reverse of the pasted input.
            foreach (string line in lines.Reverse())
            {
                var parts = SplitLine(line);
                string p5 = parts.Length > 0 ? parts[0].Trim() : string.Empty;
                string p6 = parts.Length > 1 ? parts[1].Trim() : string.Empty;

                if (string.IsNullOrEmpty(p5) && string.IsNullOrEmpty(p6))
                    continue;

                // check if row already exists (case insensitive comparison)
                if (_rows.Any(r => string.Equals(r.Col5Text, p5, StringComparison.OrdinalIgnoreCase) &&
                                   string.Equals(r.Col6Text, p6, StringComparison.OrdinalIgnoreCase)))
                    continue;

                _rows.Insert(0, new CrossRow { Col5Text = p5, Col6Text = p6, SelectedOption = 1 });
                added++;
            }
            if (added > 0)
            {
                CrossGrid.Items.Refresh();
                Save();
            }
        }

        /// <summary>
        /// Handler for the Add button.  Processes the current textbox contents and clears it.
        /// </summary>
        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            AddFromRawText(PasteTextBox.Text);
            PasteTextBox.Text = string.Empty;
        }

        /// <summary>
        /// Handler for the Clear button.  Removes all defined mappings.
        /// </summary>
        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            _rows.Clear();
            CrossGrid.Items.Refresh();
            Save();
        }

        /// <summary>
        /// Removes the currently selected rows from the grid.  Prompts the user if nothing is selected.
        /// </summary>
        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            var selected = CrossGrid.SelectedItems.Cast<CrossRow>().ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show("Please select at least one row to remove.", "No Selection", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            foreach (var r in selected)
            {
                _rows.Remove(r);
            }
            CrossGrid.Items.Refresh();
            Save();
        }

        /// <summary>
        /// Applies the current mappings via the callback and closes the window.  A copy of the list
        /// is passed back to avoid exposing the internal collection.
        /// </summary>
        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            Save();
            _onApply(_rows.ToList());
            Close();
        }

        /// <summary>
        /// Cancels the dialog without applying changes.
        /// </summary>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Handles keyboard shortcuts at window level for number keys 1-4
        /// </summary>
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            // Only handle if DataGrid has focus or selected items
            if (CrossGrid.SelectedItems.Count > 0)
            {
                if (e.Key == Key.D1 || e.Key == Key.NumPad1)
                {
                    SetGaugeForSelection(1);
                    e.Handled = true;
                }
                else if (e.Key == Key.D2 || e.Key == Key.NumPad2)
                {
                    SetGaugeForSelection(2);
                    e.Handled = true;
                }
                else if (e.Key == Key.D3 || e.Key == Key.NumPad3)
                {
                    SetGaugeForSelection(3);
                    e.Handled = true;
                }
                else if (e.Key == Key.D4 || e.Key == Key.NumPad4)
                {
                    SetGaugeForSelection(4);
                    e.Handled = true;
                }
            }
        }

        /// <summary>
        /// Handles key presses on the DataGrid to enable setting tverrsnitt options via 1..4 keys.
        /// VIKTIG: Forhindrer piltaster fra å forlate DataGrid når vi er på første/siste rad.
        /// </summary>
        private void CrossGrid_KeyDown(object sender, KeyEventArgs e)
        {
            // Handle tall-taster
            if (e.Key == Key.D1 || e.Key == Key.NumPad1)
            {
                SetGaugeForSelection(1);
                e.Handled = true;
                return;
            }
            else if (e.Key == Key.D2 || e.Key == Key.NumPad2)
            {
                SetGaugeForSelection(2);
                e.Handled = true;
                return;
            }
            else if (e.Key == Key.D3 || e.Key == Key.NumPad3)
            {
                SetGaugeForSelection(3);
                e.Handled = true;
                return;
            }
            else if (e.Key == Key.D4 || e.Key == Key.NumPad4)
            {
                SetGaugeForSelection(4);
                e.Handled = true;
                return;
            }

            // VIKTIG: Forhindre at piltaster forlater DataGrid
            if (e.Key == Key.Up || e.Key == Key.Down)
            {
                var currentItem = CrossGrid.CurrentItem;
                if (currentItem == null) return;

                int currentIndex = CrossGrid.Items.IndexOf(currentItem);

                // Hvis vi er på første rad og trykker UP - stopp!
                if (e.Key == Key.Up && currentIndex == 0)
                {
                    e.Handled = true;
                    return;
                }

                // Hvis vi er på siste rad og trykker DOWN - stopp!
                if (e.Key == Key.Down && currentIndex == CrossGrid.Items.Count - 1)
                {
                    e.Handled = true;
                    return;
                }
            }
        }

        /// <summary>
        /// Assigns the specified tverrsnitt option to all selected rows.  Refreshes the grid and saves
        /// the state afterwards. VIKTIG: Bevarer fokus og seleksjon slik at brukeren kan fortsette å navigere.
        /// </summary>
        /// <param name="option">The option index (1=1.0, 2=1.5, 3=2.5, 4=4.0).</param>
        private void SetGaugeForSelection(int option)
        {
            if (CrossGrid.SelectedItems.Count == 0) return;

            // Lagre hvilke rader som var valgt
            var selectedRows = CrossGrid.SelectedItems.Cast<CrossRow>().ToList();
            var currentItemIndex = CrossGrid.Items.IndexOf(CrossGrid.CurrentItem);

            // Oppdater verdiene
            foreach (var row in selectedRows)
            {
                row.SelectedOption = option;
            }

            // Refresh grid
            CrossGrid.Items.Refresh();

            // Gjenopprett seleksjon og fokus
            CrossGrid.SelectedItems.Clear();
            foreach (var row in selectedRows)
            {
                CrossGrid.SelectedItems.Add(row);
            }

            // Sett fokus tilbake til samme rad
            if (currentItemIndex >= 0 && currentItemIndex < CrossGrid.Items.Count)
            {
                CrossGrid.CurrentItem = CrossGrid.Items[currentItemIndex];
                CrossGrid.ScrollIntoView(CrossGrid.Items[currentItemIndex]);
            }

            // VIKTIG: Sett fokus tilbake til DataGrid slik at piltaster fungerer
            CrossGrid.Focus();

            Save();
        }

        // Context menu handlers simply forward to SetGaugeForSelection
        private void SetGauge1_Click(object sender, RoutedEventArgs e) => SetGaugeForSelection(1);
        private void SetGauge15_Click(object sender, RoutedEventArgs e) => SetGaugeForSelection(2);
        private void SetGauge25_Click(object sender, RoutedEventArgs e) => SetGaugeForSelection(3);
        private void SetGauge40_Click(object sender, RoutedEventArgs e) => SetGaugeForSelection(4);
    }
}