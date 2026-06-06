using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows;

namespace CoffeeAutoButton
{
    public static class PresetManager
    {
        public static List<Preset> LoadPresets()
        {
            AppPaths.MigrateLegacyFile("presets.json", AppPaths.PresetsPath);
            if (!File.Exists(AppPaths.PresetsPath))
            {
                return new List<Preset>();
            }

            try
            {
                var jsonString = File.ReadAllText(AppPaths.PresetsPath);
                var presets = JsonSerializer.Deserialize<List<Preset>>(jsonString) ?? new List<Preset>();
                presets.ForEach(Normalize);
                return presets;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"プリセットの読み込みに失敗しました: {ex.Message}", "エラー");
                return new List<Preset>();
            }
        }

        public static void SavePresets(List<Preset> presets)
        {
            try
            {
                AppPaths.EnsureCreated();
                presets.ForEach(Normalize);
                var jsonString = JsonSerializer.Serialize(presets, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(AppPaths.PresetsPath, jsonString);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"プリセットの保存に失敗しました: {ex.Message}", "エラー");
            }
        }

        private static void Normalize(Preset preset)
        {
            var previousVersion = preset.SchemaVersion;
            preset.SchemaVersion = AppSchema.CurrentVersion;
            if (previousVersion < 2)
            {
                preset.AutoFallbackToPhysicalClick = false;
            }
            preset.Name = string.IsNullOrWhiteSpace(preset.Name) ? "名称なし" : preset.Name;
            preset.HoldDurationText = string.IsNullOrWhiteSpace(preset.HoldDurationText) ? "500" : preset.HoldDurationText;
            preset.KeyHoldDurationText = string.IsNullOrWhiteSpace(preset.KeyHoldDurationText) ? "500" : preset.KeyHoldDurationText;
            preset.KeySequenceText ??= string.Empty;
            preset.TargetWindowTitle ??= string.Empty;
            preset.TargetWindowClass ??= string.Empty;
            preset.TargetProcessName ??= string.Empty;
            preset.DedicatedBrowserUrl ??= string.Empty;
            preset.IntervalText = string.IsNullOrWhiteSpace(preset.IntervalText) ? "10000" : preset.IntervalText;
            preset.DurationText = string.IsNullOrWhiteSpace(preset.DurationText) ? "0" : preset.DurationText;
            preset.StartDelayText = string.IsNullOrWhiteSpace(preset.StartDelayText) ? "0" : preset.StartDelayText;
        }
    }
}
