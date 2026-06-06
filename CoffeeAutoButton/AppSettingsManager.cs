using System;
using System.IO;
using System.Text.Json;
using System.Windows;

namespace CoffeeAutoButton
{
    public static class AppSettingsManager
    {
        public static AppSettings Load()
        {
            AppPaths.MigrateLegacyFile("last-settings.json", AppPaths.LastSettingsPath);
            if (!File.Exists(AppPaths.LastSettingsPath))
            {
                return new AppSettings();
            }

            try
            {
                var json = File.ReadAllText(AppPaths.LastSettingsPath);
                var settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                Normalize(settings);
                return settings;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"前回設定の読み込みに失敗しました: {ex.Message}", "エラー");
                return new AppSettings();
            }
        }

        public static void Save(AppSettings settings)
        {
            try
            {
                AppPaths.EnsureCreated();
                Normalize(settings);
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(AppPaths.LastSettingsPath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"前回設定の保存に失敗しました: {ex.Message}", "エラー");
            }
        }

        private static void Normalize(AppSettings settings)
        {
            var previousVersion = settings.SchemaVersion;
            settings.SchemaVersion = AppSchema.CurrentVersion;
            if (previousVersion < 2)
            {
                settings.AutoFallbackToPhysicalClick = false;
            }

            settings.HoldDurationText = string.IsNullOrWhiteSpace(settings.HoldDurationText) ? "500" : settings.HoldDurationText;
            settings.KeyHoldDurationText = string.IsNullOrWhiteSpace(settings.KeyHoldDurationText) ? "500" : settings.KeyHoldDurationText;
            settings.KeySequenceText ??= string.Empty;
            settings.TargetWindowTitle ??= string.Empty;
            settings.TargetWindowClass ??= string.Empty;
            settings.TargetProcessName ??= string.Empty;
            settings.DedicatedBrowserUrl ??= string.Empty;
            settings.IntervalText = string.IsNullOrWhiteSpace(settings.IntervalText) ? "10000" : settings.IntervalText;
            settings.DurationText = string.IsNullOrWhiteSpace(settings.DurationText) ? "0" : settings.DurationText;
            settings.StartDelayText = string.IsNullOrWhiteSpace(settings.StartDelayText) ? "0" : settings.StartDelayText;
            settings.StopHotkeyModifiers = settings.StopHotkeyModifiers == 0 ? HotkeyDefaults.StopModifiers : settings.StopHotkeyModifiers;
            settings.StopHotkeyKey = settings.StopHotkeyKey == 0 ? HotkeyDefaults.StopKey : settings.StopHotkeyKey;
            settings.PauseHotkeyModifiers = settings.PauseHotkeyModifiers == 0 ? HotkeyDefaults.PauseModifiers : settings.PauseHotkeyModifiers;
            settings.PauseHotkeyKey = settings.PauseHotkeyKey == 0 ? HotkeyDefaults.PauseKey : settings.PauseHotkeyKey;
        }
    }
}
