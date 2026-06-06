using System;
using System.IO;

namespace CoffeeAutoButton
{
    public static class AppPaths
    {
        private const string AppFolderName = "CoffeeAutoButton";

        public static string RootPath { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            AppFolderName);

        public static string PresetsPath => Path.Combine(RootPath, "presets.json");

        public static string LastSettingsPath => Path.Combine(RootPath, "last-settings.json");

        public static void EnsureCreated()
        {
            Directory.CreateDirectory(RootPath);
        }

        public static void MigrateLegacyFile(string legacyFileName, string targetPath)
        {
            if (File.Exists(targetPath))
            {
                return;
            }

            var legacyPaths = new[]
            {
                Path.Combine(Environment.CurrentDirectory, legacyFileName),
                Path.Combine(AppContext.BaseDirectory, legacyFileName)
            };

            foreach (var legacyPath in legacyPaths)
            {
                if (!File.Exists(legacyPath)
                    || string.Equals(Path.GetFullPath(legacyPath), Path.GetFullPath(targetPath), StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                EnsureCreated();
                File.Copy(legacyPath, targetPath);
                return;
            }
        }
    }
}
