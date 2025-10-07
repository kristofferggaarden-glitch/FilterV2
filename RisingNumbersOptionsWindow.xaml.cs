using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;

namespace FilterV1
{
    /// <summary>
    /// Interaction logic for RisingNumbersOptionsWindow.xaml
    /// Provides a UI for specifying exceptions when removing rising number pairs.
    /// All text entries are persisted between sessions, regardless of their checked state.
    /// </summary>
    public partial class RisingNumbersOptionsWindow : Window
    {
        // Klasse for å lagre både tekst og checked status
        public class ExceptionItem
        {
            public string Text { get; set; } = "";
            public bool IsChecked { get; set; } = true;
        }

        private readonly Action<List<string>> _callback;
        private readonly List<ExceptionItem> _allExceptions;
        private readonly string _settingsFilePath;

        public RisingNumbersOptionsWindow(List<string> currentlySelectedExceptions, Action<List<string>> callback)
        {
            InitializeComponent();
            _callback = callback;

            // Sett opp filbane for lagring
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appFolder = Path.Combine(appDataPath, "FilterV1");
            Directory.CreateDirectory(appFolder);
            _settingsFilePath = Path.Combine(appFolder, "RisingNumbersExceptions.json");

            // Last alle lagrede unntak
            _allExceptions = LoadAllExceptions();

            // Oppdater checked status basert på hva som ble sendt inn
            UpdateCheckedStatus(currentlySelectedExceptions ?? new List<string>());

            PopulateListBox();
        }

        private List<ExceptionItem> LoadAllExceptions()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    string json = File.ReadAllText(_settingsFilePath);
                    var loaded = JsonSerializer.Deserialize<List<ExceptionItem>>(json);
                    return loaded ?? new List<ExceptionItem>();
                }
            }
            catch (Exception)
            {
                // Ignorer feil ved lasting, bruk tom liste
            }

            return new List<ExceptionItem>();
        }

        private void SaveAllExceptions()
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(_allExceptions, options);
                File.WriteAllText(_settingsFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Kunne ikke lagre unntak: {ex.Message}", "Lagringsfeil",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void UpdateCheckedStatus(List<string> selectedExceptions)
        {
            // Oppdater checked status basert på hva som er valgt
            foreach (var exception in _allExceptions)
            {
                exception.IsChecked = selectedExceptions.Any(s =>
                    s.Equals(exception.Text, StringComparison.OrdinalIgnoreCase));
            }

            // Legg til nye unntak som ikke finnes fra før (fra tidligere økter uten persistent lagring)
            foreach (var selected in selectedExceptions)
            {
                if (!_allExceptions.Any(e => e.Text.Equals(selected, StringComparison.OrdinalIgnoreCase)))
                {
                    _allExceptions.Add(new ExceptionItem { Text = selected, IsChecked = true });
                }
            }
        }

        private void PopulateListBox()
        {
            ExceptionsListBox.Items.Clear();

            // Sorter alfabetisk for bedre oversikt
            var sortedExceptions = _allExceptions.OrderBy(e => e.Text).ToList();

            foreach (var exception in sortedExceptions)
            {
                var cb = new CheckBox
                {
                    Content = exception.Text,
                    IsChecked = exception.IsChecked,
                    Tag = exception,
                    Margin = new Thickness(5)
                };

                // Event handler for å oppdatere status når brukeren endrer
                cb.Checked += (s, e) => { exception.IsChecked = true; SaveAllExceptions(); };
                cb.Unchecked += (s, e) => { exception.IsChecked = false; SaveAllExceptions(); };

                ExceptionsListBox.Items.Add(cb);
            }
        }

        private void AddExceptionButton_Click(object sender, RoutedEventArgs e)
        {
            var txt = NewExceptionTextBox.Text.Trim();
            if (string.IsNullOrEmpty(txt)) return;

            // Sjekk om unntaket allerede finnes
            var existing = _allExceptions.FirstOrDefault(ex =>
                ex.Text.Equals(txt, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                // Hvis det finnes, huk det av og oppdater UI
                existing.IsChecked = true;
                foreach (CheckBox cb in ExceptionsListBox.Items)
                {
                    if (cb.Tag == existing)
                    {
                        cb.IsChecked = true;
                        break;
                    }
                }
            }
            else
            {
                // Legg til nytt unntak - DEFAULT UNCHECKED for sikkerhet
                var newException = new ExceptionItem { Text = txt, IsChecked = false };
                _allExceptions.Add(newException);

                // Opprett ny checkbox
                var cb = new CheckBox
                {
                    Content = txt,
                    IsChecked = false, // Default unchecked
                    Tag = newException,
                    Margin = new Thickness(5)
                };

                // Event handlers
                cb.Checked += (s, ev) => { newException.IsChecked = true; SaveAllExceptions(); };
                cb.Unchecked += (s, ev) => { newException.IsChecked = false; SaveAllExceptions(); };

                ExceptionsListBox.Items.Add(cb);

                // Sorter listen på nytt
                PopulateListBox();
            }

            SaveAllExceptions();
            NewExceptionTextBox.Clear();
            NewExceptionTextBox.Focus();
        }

        private void RemoveExceptionButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedCheckBoxes = ExceptionsListBox.SelectedItems.Cast<CheckBox>().ToList();

            if (selectedCheckBoxes.Count == 0)
            {
                MessageBox.Show("Velg ett eller flere unntak å slette.", "Ingen valg",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string message = selectedCheckBoxes.Count == 1
                ? "Er du sikker på at du vil slette det valgte unntaket?"
                : $"Er du sikker på at du vil slette {selectedCheckBoxes.Count} unntak?";

            if (MessageBox.Show(message, "Bekreft sletting", MessageBoxButton.YesNo,
                MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                foreach (var cb in selectedCheckBoxes)
                {
                    if (cb.Tag is ExceptionItem exception)
                    {
                        _allExceptions.Remove(exception);
                    }
                }

                SaveAllExceptions();
                PopulateListBox();
            }
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (CheckBox cb in ExceptionsListBox.Items)
            {
                cb.IsChecked = true;
                if (cb.Tag is ExceptionItem exception)
                {
                    exception.IsChecked = true;
                }
            }
            SaveAllExceptions();
        }

        private void DeselectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (CheckBox cb in ExceptionsListBox.Items)
            {
                cb.IsChecked = false;
                if (cb.Tag is ExceptionItem exception)
                {
                    exception.IsChecked = false;
                }
            }
            SaveAllExceptions();
        }

        private void NewExceptionTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                AddExceptionButton_Click(sender, e);
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // Returner bare de som er huket av
            var selectedExceptions = _allExceptions
                .Where(ex => ex.IsChecked)
                .Select(ex => ex.Text)
                .ToList();

            SaveAllExceptions(); // Sørg for at alt er lagret
            _callback?.Invoke(selectedExceptions);
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            // Ved avbryt, returner de som var huket av når vinduet åpnet
            var originallySelected = _allExceptions
                .Where(ex => ex.IsChecked)
                .Select(ex => ex.Text)
                .ToList();

            _callback?.Invoke(originallySelected);
            this.Close();
        }
    }
}