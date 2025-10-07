using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FilterV1
{
    public partial class MarkUnmarkedDupesWindow : Window
    {
        private Action<List<string>> _callback;
        private List<MarkDupePatternViewModel> _patterns;
        private readonly string _settingsFilePath;

        public MarkUnmarkedDupesWindow(Action<List<string>> callback)
        {
            InitializeComponent();
            _callback = callback;

            // Setup settings file path
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appFolder = Path.Combine(appDataPath, "FilterV1");
            Directory.CreateDirectory(appFolder);
            _settingsFilePath = Path.Combine(appFolder, "MarkDupePatterns.json");

            LoadPatterns();
            RefreshPatternsList();
        }

        private void LoadPatterns()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    string json = File.ReadAllText(_settingsFilePath);
                    var loaded = JsonSerializer.Deserialize<List<MarkDupePattern>>(json);
                    _patterns = loaded?.Select(p => new MarkDupePatternViewModel
                    {
                        Pattern = p.Pattern ?? "",
                        IsEnabled = p.IsEnabled
                    }).ToList() ?? new List<MarkDupePatternViewModel>();
                }
                else
                {
                    _patterns = new List<MarkDupePatternViewModel>();
                }
            }
            catch
            {
                _patterns = new List<MarkDupePatternViewModel>();
            }
        }

        private void SavePatterns()
        {
            try
            {
                var toSave = _patterns.Select(p => new MarkDupePattern
                {
                    Pattern = p.Pattern,
                    IsEnabled = p.IsEnabled
                }).ToList();

                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(toSave, options);
                File.WriteAllText(_settingsFilePath, json);
            }
            catch
            {
                // Ignore save errors
            }
        }

        private void RefreshPatternsList()
        {
            PatternsListBox.ItemsSource = null;
            PatternsListBox.ItemsSource = _patterns;
        }

        private void PatternTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddPatternButton_Click(sender, e);
            }
        }

        private void AddPatternButton_Click(object sender, RoutedEventArgs e)
        {
            string pattern = PatternTextBox.Text.Trim();
            if (string.IsNullOrEmpty(pattern))
            {
                MessageBox.Show("Vennligst skriv inn et tekstmønster.", "Valideringsfeil",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_patterns.Any(p => p.Pattern.Equals(pattern, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("Dette mønsteret finnes allerede.", "Duplikat",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _patterns.Add(new MarkDupePatternViewModel
            {
                Pattern = pattern,
                IsEnabled = true
            });

            PatternTextBox.Clear();
            PatternTextBox.Focus();
            RefreshPatternsList();
            SavePatterns();
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var pattern in _patterns)
            {
                pattern.IsEnabled = true;
            }
            PatternsListBox.Items.Refresh();
            SavePatterns();
        }

        private void DeselectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var pattern in _patterns)
            {
                pattern.IsEnabled = false;
            }
            PatternsListBox.Items.Refresh();
            SavePatterns();
        }

        private void RemovePatternButton_Click(object sender, RoutedEventArgs e)
        {
            if (PatternsListBox.SelectedIndex >= 0)
            {
                var selectedPattern = _patterns[PatternsListBox.SelectedIndex];
                _patterns.Remove(selectedPattern);
                RefreshPatternsList();
                SavePatterns();
            }
            else
            {
                MessageBox.Show("Vennligst velg et mønster å fjerne.", "Ingen valg",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ApplyMarkingButton_Click(object sender, RoutedEventArgs e)
        {
            var enabledPatterns = _patterns.Where(p => p.IsEnabled).Select(p => p.Pattern).ToList();

            if (enabledPatterns.Count == 0)
            {
                MessageBox.Show("Vennligst aktiver minst ett mønster.", "Ingen aktive mønstre",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SavePatterns();
            _callback?.Invoke(enabledPatterns);
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    public class MarkDupePattern
    {
        public string Pattern { get; set; } = "";
        public bool IsEnabled { get; set; } = true;
    }

    public class MarkDupePatternViewModel
    {
        public string Pattern { get; set; } = "";
        public bool IsEnabled { get; set; } = true;
        public string DisplayText => $"Marker celler som inneholder: '{Pattern}'";
    }
}