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
        // Holds the master list of all mappings entered by the user.  This collection contains
        // the actual CrossRow objects which will be persisted and passed back via the callback.
        private ObservableCollection<CrossRow> _rows = new ObservableCollection<CrossRow>();
        // Holds the current view of rows after search filtering.  Bound to the DataGrid.
        private ObservableCollection<CrossRow> _filteredRows = new ObservableCollection<CrossRow>();
        // Tracks the anchor index for shift+arrow multi‑selection.  When null no anchor is set.
        private int? _selectionAnchorIndex = null;
        // Stores the current search term so that new entries can be filtered immediately.
        private string _currentSearchTerm = string.Empty;

        // Stores the available cross options loaded from persistent storage.  Each option
        // defines a numeric Id, a user visible label, and three text values corresponding
        // to the columns that will be filled when the option is applied.  The list does
        // not include the blank option (Id=0) which is handled separately.  Options are
        // loaded via LoadCrossOptions() and can be edited from the UI.
        private List<CrossOption> _crossOptions = new List<CrossOption>();

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

            // Load persisted cross options from disk and update the ComboBox list.  The loaded
            // options include user defined definitions and are supplemented by a blank entry.
            LoadCrossOptions();

            // Load rows either from the provided list or from persistent storage.
            Load(existing);
            // Bind the DataGrid to the filtered collection.  This must be set before
            // invoking FilterRows so that the Items.Refresh call operates on the correct view.
            CrossGrid.ItemsSource = _filteredRows;
            // Initially populate the filtered list with all rows.  Subsequent calls to
            // FilterRows() will update this collection based on the user's search input.
            FilterRows(string.Empty);

            // Handle keyboard shortcuts at Window level
            this.PreviewKeyDown += Window_PreviewKeyDown;
        }

        /// <summary>
        /// Loads cross section options from persistent storage and rebuilds the ComboBox
        /// items list.  If no file exists or the file is invalid, default options are
        /// created and saved.  After loading, any existing SelectedOption values on
        /// rows that are no longer valid will be reset to 0.  Finally the DataGrid is
        /// refreshed to reflect updated labels.
        /// </summary>
        private void LoadCrossOptions()
        {
            _crossOptions = CrossOptionRepository.Load();
            UpdateOptionItems();
        }

        /// <summary>
        /// Updates the Tag property on the window with a list of OptionItem objects based on
        /// the currently loaded cross options.  A blank option (Value=0) is always added at
        /// the beginning of the list.  After updating the Tag, any rows referencing
        /// non‑existent options are reset to blank.  Finally the grid is refreshed.
        /// </summary>
        private void UpdateOptionItems()
        {
            var items = new List<OptionItem>();
            // Add blank/none option
            items.Add(new OptionItem { Value = 0, Label = string.Empty });
            foreach (var opt in _crossOptions.OrderBy(o => o.Id))
            {
                items.Add(new OptionItem { Value = opt.Id, Label = opt.Label });
            }
            this.Tag = items;

            // Validate selected options on existing rows
            foreach (var row in _rows)
            {
                if (row.SelectedOption > 0 && !_crossOptions.Any(o => o.Id == row.SelectedOption))
                {
                    row.SelectedOption = 0;
                }
            }
            CrossGrid?.Items.Refresh();
        }

        /// <summary>
        /// Handles click on the edit options button.  Opens the CrossOptionSettingsWindow and
        /// updates the internal list when the user saves changes.  After updating, the
        /// ComboBox items are refreshed.  Selected options on rows that no longer exist
        /// are automatically reset to blank.
        /// </summary>
        private void EditOptionsButton_Click(object sender, RoutedEventArgs e)
        {
            var optWindow = new CrossOptionSettingsWindow(_crossOptions, updatedOptions =>
            {
                // Update internal list and refresh UI
                _crossOptions = updatedOptions;
                UpdateOptionItems();
            });
            optWindow.Owner = this;
            optWindow.ShowDialog();
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
                        // Validate SelectedOption against the currently loaded cross options.  If the
                        // stored value is 0 (blank) or matches an existing option Id, retain it;
                        // otherwise reset to 0 so that invalid references do not persist.
                        SelectedOption = (r.SelectedOption >= 0 &&
                                          (r.SelectedOption == 0 || _crossOptions.Any(o => o.Id == r.SelectedOption)))
                            ? r.SelectedOption
                            : 0
                    });
                }
            }
            catch
            {
                _rows.Clear();
            }

            // Whenever rows are loaded from disk or provided externally, refresh the filtered view
            // to include all rows by default.  If a search term is currently active it will be
            // applied in FilterRows().  This call also clears any existing selections and anchors.
            FilterRows(_currentSearchTerm);
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
                // Only copy data from column 5 and 6 when pasting from the data preview.
                // If the pasted line contains six or more columns (tab/semicolon/comma separated), then
                // assume the first four are columns 1–4 and use elements at index 4 and 5 for col5 and col6.
                // For lines with fewer than six parts, treat the first two values as columns 5 and 6.
                string raw5 = string.Empty;
                string raw6 = string.Empty;
                if (parts.Length >= 6)
                {
                    // When there are six or more parts, the input likely contains columns 1–6 (and beyond).
                    // Use the 5th (index 4) and 6th (index 5) values as column 5 and column 6.
                    raw5 = parts[4].Trim();
                    raw6 = parts[5].Trim();
                }
                else if (parts.Length >= 2)
                {
                    // For lines with fewer than six parts (for example when the user copies only columns 5 & 6,
                    // or columns 5–7), assume the first two entries are the values from columns 5 and 6.  This
                    // avoids accidentally pulling in values from later columns (e.g. column 7) when pasting.
                    raw5 = parts[0].Trim();
                    raw6 = parts[1].Trim();
                }
                else if (parts.Length == 1)
                {
                    raw5 = parts[0].Trim();
                    raw6 = string.Empty;
                }

                if (string.IsNullOrEmpty(raw5) && string.IsNullOrEmpty(raw6))
                    continue;

                // Strip prefix before the first dash '-' to remove device/prefix identifiers like J02-
                string Sanitize(string val)
                {
                    if (string.IsNullOrWhiteSpace(val)) return string.Empty;
                    int dash = val.IndexOf('-');
                    return dash > 0 ? val.Substring(dash + 1).Trim() : val.Trim();
                }
                string p5 = Sanitize(raw5);
                string p6 = Sanitize(raw6);

                // Skip if both sanitized values are empty
                if (string.IsNullOrEmpty(p5) && string.IsNullOrEmpty(p6))
                    continue;

                // check if row already exists (case insensitive comparison) based on sanitized values
                if (_rows.Any(r => string.Equals(r.Col5Text, p5, StringComparison.OrdinalIgnoreCase) &&
                                   string.Equals(r.Col6Text, p6, StringComparison.OrdinalIgnoreCase)))
                    continue;

                // Insert new row at the beginning so newest appears on top.  SelectedOption defaulted to 0.
                _rows.Insert(0, new CrossRow { Col5Text = p5, Col6Text = p6, SelectedOption = 0 });
                added++;
            }
            if (added > 0)
            {
                // Refresh the filtered view based on current search term so that new rows appear appropriately
                FilterRows(_currentSearchTerm);
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
            // Clear anchor and filtered list
            _selectionAnchorIndex = null;
            FilterRows(_currentSearchTerm);
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
                // Remove from master list
                _rows.Remove(r);
            }
            // Clear anchor and update filtered view
            _selectionAnchorIndex = null;
            FilterRows(_currentSearchTerm);
            Save();
        }

        /// <summary>
        /// Applies the current mappings via the callback and closes the window.  A copy of the list
        /// is passed back to avoid exposing the internal collection.
        /// </summary>
        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            Save();
            // Pass back a copy of the master list, not the filtered view
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
            // If already handled by a child control (e.g. DataGrid preview handler), do nothing
            if (e.Handled)
            {
                return;
            }

            // Only handle if there is a selection.  Use the same mapping logic as in
            // CrossGrid_PreviewKeyDown: digits map to options; Ctrl+digits map to higher
            // options starting at 6.  0 resets to blank.  Ignore presses that refer to
            // undefined options.
            if (CrossGrid.SelectedItems.Count > 0)
            {
                bool ctrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
                int option = 0;
                switch (e.Key)
                {
                    case Key.D1:
                    case Key.NumPad1:
                        option = ctrl ? 6 : 1;
                        break;
                    case Key.D2:
                    case Key.NumPad2:
                        option = ctrl ? 7 : 2;
                        break;
                    case Key.D3:
                    case Key.NumPad3:
                        option = ctrl ? 8 : 3;
                        break;
                    case Key.D4:
                    case Key.NumPad4:
                        option = ctrl ? 9 : 4;
                        break;
                    case Key.D5:
                    case Key.NumPad5:
                        option = ctrl ? 10 : 5;
                        break;
                    case Key.D6:
                    case Key.NumPad6:
                        option = 6;
                        break;
                    case Key.D7:
                    case Key.NumPad7:
                        option = 7;
                        break;
                    case Key.D8:
                    case Key.NumPad8:
                        option = 8;
                        break;
                    case Key.D9:
                    case Key.NumPad9:
                        option = 9;
                        break;
                    case Key.D0:
                    case Key.NumPad0:
                        option = 0;
                        break;
                    default:
                        option = 0;
                        break;
                }
                if (option > 0)
                {
                    if (_crossOptions.Any(o => o.Id == option))
                    {
                        SetGaugeForSelection(option);
                        e.Handled = true;
                    }
                }
                else if (option == 0 && (e.Key == Key.D0 || e.Key == Key.NumPad0))
                {
                    SetGaugeForSelection(0);
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
            // Do not handle number keys here; they are handled in the PreviewKeyDown to ensure
            // the DataGrid doesn't interpret them as starting cell editing and moving focus.
            // Leaving this block empty allows the PreviewKeyDown handler to intercept the keys first.

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
        /// Preview handler for key presses on the DataGrid.  This runs before the DataGrid gets a chance to
        /// interpret the key and allows us to override default behavior.  We intercept number keys (1-4)
        /// regardless of modifier keys so that the DataGrid doesn't treat them as input to a cell (which
        /// otherwise triggers editing and moves focus like the Tab key).  We also handle arrow navigation
        /// with optional Shift to extend the selection and prevent leaving the DataGrid at the first/last row.
        /// </summary>
        private void CrossGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (CrossGrid == null) return;

            // If the Delete key is pressed and there are selected rows, remove them immediately
            if (e.Key == Key.Delete)
            {
                if (CrossGrid.SelectedItems.Count > 0)
                {
                    var toRemove = CrossGrid.SelectedItems.Cast<CrossRow>().ToList();
                    foreach (var row in toRemove)
                    {
                        _rows.Remove(row);
                    }
                    // Clear anchor and refresh filtered view so removed rows disappear
                    _selectionAnchorIndex = null;
                    FilterRows(_currentSearchTerm);
                    Save();
                }
                e.Handled = true;
                return;
            }

            // Determine if a shortcut key corresponds to a cross option.  Without Ctrl,
            // digits map directly to options 1..9.  With Ctrl, digits map to options
            // starting at 6 (Ctrl+1=6, Ctrl+2=7, etc.).  The 0 key resets to blank.
            int option = 0;
            bool ctrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
            switch (e.Key)
            {
                case Key.D1:
                case Key.NumPad1:
                    option = ctrl ? 6 : 1;
                    break;
                case Key.D2:
                case Key.NumPad2:
                    option = ctrl ? 7 : 2;
                    break;
                case Key.D3:
                case Key.NumPad3:
                    option = ctrl ? 8 : 3;
                    break;
                case Key.D4:
                case Key.NumPad4:
                    option = ctrl ? 9 : 4;
                    break;
                case Key.D5:
                case Key.NumPad5:
                    option = ctrl ? 10 : 5;
                    break;
                case Key.D6:
                case Key.NumPad6:
                    option = 6;
                    break;
                case Key.D7:
                case Key.NumPad7:
                    option = 7;
                    break;
                case Key.D8:
                case Key.NumPad8:
                    option = 8;
                    break;
                case Key.D9:
                case Key.NumPad9:
                    option = 9;
                    break;
                case Key.D0:
                case Key.NumPad0:
                    option = 0;
                    break;
                default:
                    option = 0;
                    break;
            }

            if (option > 0)
            {
                // Only assign if there is an active selection and the option exists
                if (CrossGrid.SelectedItems.Count > 0 && _crossOptions.Any(o => o.Id == option))
                {
                    SetGaugeForSelection(option);
                }

                // Handle the event so the DataGrid doesn't treat the key as cell input
                e.Handled = true;
                return;
            }
            else if (option == 0 && (e.Key == Key.D0 || e.Key == Key.NumPad0))
            {
                // Reset selected rows to blank
                if (CrossGrid.SelectedItems.Count > 0)
                {
                    SetGaugeForSelection(0);
                }
                e.Handled = true;
                return;
            }

            // Handle Up/Down arrow keys for navigation and multi-selection
            if (e.Key == Key.Up || e.Key == Key.Down)
            {
                // Determine current index within the filtered view
                var currentItem = CrossGrid.CurrentItem;
                int currentIndex = currentItem != null ? CrossGrid.Items.IndexOf(currentItem) : -1;

                bool shiftPressed = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
                int delta = e.Key == Key.Up ? -1 : 1;
                int targetIndex = currentIndex + delta;

                // If no current item, select the first item by default
                if (currentIndex < 0)
                {
                    if (CrossGrid.Items.Count > 0)
                    {
                        currentIndex = 0;
                        targetIndex = (e.Key == Key.Up ? 0 : 0);
                    }
                }

                // Ensure target index stays within bounds
                if (targetIndex < 0 || targetIndex >= CrossGrid.Items.Count)
                {
                    // Prevent leaving the DataGrid when at first/last row
                    e.Handled = true;
                    return;
                }

                if (shiftPressed)
                {
                    // When shift is pressed and no anchor exists, set anchor to current index
                    if (_selectionAnchorIndex == null)
                    {
                        _selectionAnchorIndex = currentIndex >= 0 ? currentIndex : targetIndex;
                    }

                    // Calculate selection range between anchor and target
                    int anchor = _selectionAnchorIndex ?? targetIndex;
                    int start = Math.Min(anchor, targetIndex);
                    int end = Math.Max(anchor, targetIndex);

                    // Clear previous selection and select the contiguous range
                    CrossGrid.SelectedItems.Clear();
                    for (int i = start; i <= end; i++)
                    {
                        var item = CrossGrid.Items[i];
                        CrossGrid.SelectedItems.Add(item);
                    }

                    // Update current item to the newly navigated row
                    var newCurrent = CrossGrid.Items[targetIndex];
                    CrossGrid.CurrentItem = newCurrent;
                    CrossGrid.ScrollIntoView(newCurrent);
                }
                else
                {
                    // Without shift, clear anchor and move selection to the target row
                    _selectionAnchorIndex = null;
                    var newCurrent = CrossGrid.Items[targetIndex];
                    CrossGrid.SelectedItems.Clear();
                    CrossGrid.SelectedItems.Add(newCurrent);
                    CrossGrid.CurrentItem = newCurrent;
                    CrossGrid.ScrollIntoView(newCurrent);
                }

                CrossGrid.Focus();
                e.Handled = true;
                return;
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

            // Capture selection and current item index based on filtered view
            var selectedRows = CrossGrid.SelectedItems.Cast<CrossRow>().ToList();
            var currentItemIndex = CrossGrid.Items.IndexOf(CrossGrid.CurrentItem);

            // Assign option to each selected row (updates underlying master list as objects are shared)
            foreach (var row in selectedRows)
            {
                row.SelectedOption = option;
            }

            // Refresh the view so that the ComboBox displays the updated option
            CrossGrid.Items.Refresh();

            // Restore selection and focus
            CrossGrid.SelectedItems.Clear();
            foreach (var row in selectedRows)
            {
                CrossGrid.SelectedItems.Add(row);
            }

            // Restore current item if still valid
            if (currentItemIndex >= 0 && currentItemIndex < CrossGrid.Items.Count)
            {
                CrossGrid.CurrentItem = CrossGrid.Items[currentItemIndex];
                CrossGrid.ScrollIntoView(CrossGrid.Items[currentItemIndex]);
            }

            CrossGrid.Focus();
            Save();
        }

        // Context menu handlers simply forward to SetGaugeForSelection
        private void SetGauge1_Click(object sender, RoutedEventArgs e) => SetGaugeForSelection(1);
        private void SetGauge15_Click(object sender, RoutedEventArgs e) => SetGaugeForSelection(2);
        private void SetGauge25_Click(object sender, RoutedEventArgs e) => SetGaugeForSelection(3);
        private void SetGauge40_Click(object sender, RoutedEventArgs e) => SetGaugeForSelection(4);

        /// <summary>
        /// Called whenever the text in the search box changes.  Updates the filtered view
        /// of mappings based on the search term.  Searching is case‑insensitive and
        /// matches against both column values.  If the search box is cleared all rows
        /// are displayed.
        /// </summary>
        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _currentSearchTerm = (sender as TextBox)?.Text ?? string.Empty;
            FilterRows(_currentSearchTerm);
        }

        /// <summary>
        /// Filters the master list (_rows) into the filtered view (_filteredRows) based on
        /// the provided search term.  Existing selections and the selection anchor are
        /// cleared to avoid mismatches between the view and underlying data.  After
        /// filtering the DataGrid is refreshed.
        /// </summary>
        /// <param name="searchTerm">The search term entered by the user.</param>
        private void FilterRows(string searchTerm)
        {
            // Clear anchor whenever filtering changes the row order/contents
            _selectionAnchorIndex = null;
            _filteredRows.Clear();

            if (string.IsNullOrWhiteSpace(searchTerm))
            {
                // When search is empty include all rows in the same order
                foreach (var row in _rows)
                {
                    _filteredRows.Add(row);
                }
            }
            else
            {
                string term = searchTerm.Trim();
                foreach (var row in _rows)
                {
                    if ((row.Col5Text != null && row.Col5Text.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0) ||
                        (row.Col6Text != null && row.Col6Text.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        _filteredRows.Add(row);
                    }
                }
            }
            CrossGrid.Items.Refresh();
        }
    }
}