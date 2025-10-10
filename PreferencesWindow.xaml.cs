using System;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;

namespace FilterV1
{
    public partial class PreferencesWindow : Window
    {
        private Action<AppPreferences> _callback;
        private AppPreferences _preferences;

        public PreferencesWindow(AppPreferences currentPrefs, Action<AppPreferences> callback)
        {
            InitializeComponent();
            _callback = callback;
            _preferences = new AppPreferences
            {
                AccentColor = currentPrefs.AccentColor,
                SuccessColor = currentPrefs.SuccessColor,
                DangerColor = currentPrefs.DangerColor,
                BackgroundColor = currentPrefs.BackgroundColor,
                PanelColor = currentPrefs.PanelColor,
                TextColor = currentPrefs.TextColor,
                IsDarkMode = currentPrefs.IsDarkMode,
                DataGridFontSize = currentPrefs.DataGridFontSize
            };

            // Set initial values
            FontSizeSlider.Value = _preferences.DataGridFontSize;
            FontSizeSlider.ValueChanged += FontSizeSlider_ValueChanged;
            UpdateFontSizeDisplay();
            UpdatePreview();
        }

        private void FontSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            _preferences.DataGridFontSize = (int)FontSizeSlider.Value;
            UpdateFontSizeDisplay();
            UpdatePreview();
        }

        private void UpdateFontSizeDisplay()
        {
            if (FontSizeDisplay != null)
            {
                FontSizeDisplay.Text = $"{_preferences.DataGridFontSize} px";
            }
        }

        private void UpdatePreview()
        {
            if (PreviewAccentText != null && PreviewDataText != null)
            {
                var accentBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(_preferences.AccentColor));
                var textBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(_preferences.TextColor));
                PreviewAccentText.Foreground = accentBrush;
                PreviewDataText.Foreground = textBrush;
                PreviewDataText.FontSize = _preferences.DataGridFontSize;
            }
        }

        private void DarkModeButton_Click(object sender, RoutedEventArgs e)
        {
            _preferences.IsDarkMode = true;
            _preferences.BackgroundColor = "#1a1a2e";
            _preferences.PanelColor = "#16213e";
            _preferences.TextColor = "#FFFFFF";
            UpdatePreview();
        }

        private void LightModeButton_Click(object sender, RoutedEventArgs e)
        {
            _preferences.IsDarkMode = false;
            _preferences.BackgroundColor = "#f5f5f5";
            _preferences.PanelColor = "#ffffff";
            _preferences.TextColor = "#2c3e50";
            UpdatePreview();
        }

        private void Accent1Button_Click(object sender, RoutedEventArgs e)
        {
            _preferences.AccentColor = "#00d4ff";
            _preferences.SuccessColor = "#00ff88";
            _preferences.DangerColor = "#ff4757";
            UpdatePreview();
        }

        private void Accent2Button_Click(object sender, RoutedEventArgs e)
        {
            _preferences.AccentColor = "#a855f7";
            _preferences.SuccessColor = "#22c55e";
            _preferences.DangerColor = "#ef4444";
            UpdatePreview();
        }

        private void Accent3Button_Click(object sender, RoutedEventArgs e)
        {
            _preferences.AccentColor = "#10b981";
            _preferences.SuccessColor = "#34d399";
            _preferences.DangerColor = "#f87171";
            UpdatePreview();
        }

        private void Accent4Button_Click(object sender, RoutedEventArgs e)
        {
            _preferences.AccentColor = "#f97316";
            _preferences.SuccessColor = "#10b981";
            _preferences.DangerColor = "#dc2626";
            UpdatePreview();
        }

        private void Accent5Button_Click(object sender, RoutedEventArgs e)
        {
            _preferences.AccentColor = "#ec4899";
            _preferences.SuccessColor = "#22d3ee";
            _preferences.DangerColor = "#fb7185";
            UpdatePreview();
        }

        private void Accent6Button_Click(object sender, RoutedEventArgs e)
        {
            _preferences.AccentColor = "#3b82f6";
            _preferences.SuccessColor = "#22c55e";
            _preferences.DangerColor = "#ef4444";
            UpdatePreview();
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "Er du sikker på at du vil tilbakestille til standard innstillinger?",
                "Tilbakestill innstillinger",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _preferences = AppPreferences.GetDefault();
                FontSizeSlider.Value = _preferences.DataGridFontSize;
                UpdatePreview();
                MessageBox.Show("Innstillinger tilbakestilt til standard.", "Tilbakestilt",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            _callback?.Invoke(_preferences);
            MessageBox.Show("Innstillinger lagret! Endringene er nå aktive.", "Lagret",
                MessageBoxButton.OK, MessageBoxImage.Information);
            // IKKE lukk vinduet her
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    public class AppPreferences
    {
        public string AccentColor { get; set; } = "#00d4ff";
        public string SuccessColor { get; set; } = "#00ff88";
        public string DangerColor { get; set; } = "#ff4757";
        public string BackgroundColor { get; set; } = "#1a1a2e";
        public string PanelColor { get; set; } = "#16213e";
        public string TextColor { get; set; } = "#FFFFFF";
        public bool IsDarkMode { get; set; } = true;
        public int DataGridFontSize { get; set; } = 13;

        public static AppPreferences GetDefault()
        {
            return new AppPreferences
            {
                AccentColor = "#00d4ff",
                SuccessColor = "#00ff88",
                DangerColor = "#ff4757",
                BackgroundColor = "#1a1a2e",
                PanelColor = "#16213e",
                TextColor = "#FFFFFF",
                IsDarkMode = true,
                DataGridFontSize = 13
            };
        }

        public static AppPreferences Load(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    return JsonSerializer.Deserialize<AppPreferences>(json) ?? GetDefault();
                }
            }
            catch
            {
                // If load fails, return defaults
            }
            return GetDefault();
        }

        public void Save(string filePath)
        {
            try
            {
                string directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(this, options);
                File.WriteAllText(filePath, json);
            }
            catch
            {
                // Ignore save errors
            }
        }
    }
}