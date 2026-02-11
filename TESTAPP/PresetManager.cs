using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows; // For MessageBox

namespace AutoKeyInputApp
{
    public static class PresetManager
    {
        private static readonly string FilePath = "presets.json";

        public static List<Preset> LoadPresets()
        {
            if (!File.Exists(FilePath))
            {
                return new List<Preset>();
            }

            try
            {
                string jsonString = File.ReadAllText(FilePath);
                return JsonSerializer.Deserialize<List<Preset>>(jsonString) ?? new List<Preset>();
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
                string jsonString = JsonSerializer.Serialize(presets, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(FilePath, jsonString);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"プリセットの保存に失敗しました: {ex.Message}", "エラー");
            }
        }
    }
}
