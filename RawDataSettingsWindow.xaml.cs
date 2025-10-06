using Microsoft.Win32;
using System;
using System.Windows;
using Ookii.Dialogs.Wpf;

namespace FilterV1
{
    public partial class RawDataSettingsWindow : Window
    {
        private readonly Action<RawDataSettings> _callback;
        private RawDataSettings _settings;

        public RawDataSettingsWindow(RawDataSettings currentSettings, Action<RawDataSettings> callback)
        {
            InitializeComponent();
            _callback = callback;
            _settings = new RawDataSettings
            {
                RawFileLocation = currentSettings.RawFileLocation,
                TemplateFile1 = currentSettings.TemplateFile1,
                TemplateFile2 = currentSettings.TemplateFile2,
                TemplateFile3 = currentSettings.TemplateFile3
            };

            UpdateUI();
        }

        private void UpdateUI()
        {
            RawFileLocationTextBox.Text = string.IsNullOrEmpty(_settings.RawFileLocation)
                ? "Ikke valgt"
                : _settings.RawFileLocation;

            TemplateFile1TextBox.Text = string.IsNullOrEmpty(_settings.TemplateFile1)
                ? "Ikke valgt"
                : _settings.TemplateFile1;

            TemplateFile2TextBox.Text = string.IsNullOrEmpty(_settings.TemplateFile2)
                ? "Ikke valgt"
                : _settings.TemplateFile2;

            TemplateFile3TextBox.Text = string.IsNullOrEmpty(_settings.TemplateFile3)
                ? "Ikke valgt"
                : _settings.TemplateFile3;
        }

        private void SelectRawFileLocationButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new VistaFolderBrowserDialog
            {
                Description = "Velg mappe hvor råfilene ligger",
                UseDescriptionForTitle = true,
                SelectedPath = _settings.RawFileLocation
            };

            if (dialog.ShowDialog() == true)
            {
                _settings.RawFileLocation = dialog.SelectedPath;
                UpdateUI();
            }
        }

        private void SelectTemplateFile1Button_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Velg malfil 1 (mottar data, suffiks F)",
                Filter = "Excel Files|*.xlsx;*.xls",
                InitialDirectory = string.IsNullOrEmpty(_settings.TemplateFile1)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    : System.IO.Path.GetDirectoryName(_settings.TemplateFile1)
            };

            if (dialog.ShowDialog() == true)
            {
                _settings.TemplateFile1 = dialog.FileName;
                UpdateUI();
            }
        }

        private void SelectTemplateFile2Button_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Velg malfil 2 (Nord, suffiks N)",
                Filter = "Excel Files|*.xlsx;*.xls",
                InitialDirectory = string.IsNullOrEmpty(_settings.TemplateFile2)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    : System.IO.Path.GetDirectoryName(_settings.TemplateFile2)
            };

            if (dialog.ShowDialog() == true)
            {
                _settings.TemplateFile2 = dialog.FileName;
                UpdateUI();
            }
        }

        private void SelectTemplateFile3Button_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Velg malfil 3 (Durapart, suffiks D)",
                Filter = "Excel Files|*.xlsx;*.xls",
                InitialDirectory = string.IsNullOrEmpty(_settings.TemplateFile3)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    : System.IO.Path.GetDirectoryName(_settings.TemplateFile3)
            };

            if (dialog.ShowDialog() == true)
            {
                _settings.TemplateFile3 = dialog.FileName;
                UpdateUI();
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Validation
            if (string.IsNullOrEmpty(_settings.RawFileLocation))
            {
                MessageBox.Show("Vennligst velg en råfil-lokasjon.", "Valideringsfeil",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(_settings.TemplateFile1) ||
                string.IsNullOrEmpty(_settings.TemplateFile2) ||
                string.IsNullOrEmpty(_settings.TemplateFile3))
            {
                MessageBox.Show("Vennligst velg alle tre malfilene.", "Valideringsfeil",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _callback?.Invoke(_settings);
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}