using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace FilterV1
{
    /// <summary>
    /// Interaction logic for CrossOptionSettingsWindow.xaml
    ///
    /// This window allows the user to edit, add and remove custom cross section options.  Each option
    /// maps a numeric shortcut (Id) to a display label and three column values.  The list of
    /// options is persisted via CrossOptionRepository.  When saved, the provided callback
    /// returns a new list which the caller can apply.
    /// </summary>
    public partial class CrossOptionSettingsWindow : Window
    {
        private ObservableCollection<CrossOption> _options;
        private readonly Action<List<CrossOption>> _onSaved;

        public CrossOptionSettingsWindow(List<CrossOption> currentOptions, Action<List<CrossOption>> onSaved)
        {
            InitializeComponent();
            _onSaved = onSaved;

            // Create a deep copy of the options to edit locally.  This prevents direct
            // modification of the caller's list until Save is invoked.
            _options = new ObservableCollection<CrossOption>(
                currentOptions.Select(opt => new CrossOption
                {
                    Id = opt.Id,
                    Label = opt.Label,
                    Col2 = opt.Col2,
                    Col3 = opt.Col3,
                    Col4 = opt.Col4
                })
            );
            OptionsGrid.ItemsSource = _options;
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            // Determine next available Id (max existing + 1).  Ensure Id is not zero.
            int nextId = 1;
            if (_options.Count > 0)
            {
                nextId = _options.Max(o => o.Id) + 1;
            }
            _options.Add(new CrossOption { Id = nextId, Label = string.Empty, Col2 = string.Empty, Col3 = string.Empty, Col4 = string.Empty });
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            var selected = OptionsGrid.SelectedItems.Cast<CrossOption>().ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show("Velg minst én rad å fjerne.", "Ingen markering", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            foreach (var opt in selected)
            {
                _options.Remove(opt);
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Create a normalized list: trim strings and ensure Ids are positive
            var finalList = _options
                .Where(o => o != null && o.Id > 0)
                .Select(o => new CrossOption
                {
                    Id = o.Id,
                    Label = o.Label?.Trim() ?? string.Empty,
                    Col2 = o.Col2?.Trim() ?? string.Empty,
                    Col3 = o.Col3?.Trim() ?? string.Empty,
                    Col4 = o.Col4?.Trim() ?? string.Empty
                })
                .OrderBy(o => o.Id)
                .ToList();

            // Invoke callback before saving to allow the caller to update its state
            _onSaved?.Invoke(finalList);

            // Save to repository for persistence
            CrossOptionRepository.Save(finalList);

            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}