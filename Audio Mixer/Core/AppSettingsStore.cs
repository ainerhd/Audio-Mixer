using System;
using System.IO;
using System.Text.Json;

namespace Audio_Mixer.Core
{
    public sealed class AppSettingsStore
    {
        private readonly string settingsPath;

        public AppSettingsStore()
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Audio_Mixer");
            Directory.CreateDirectory(directory);
            settingsPath = Path.Combine(directory, "app_settings.json");
        }

        public AppSettings Load()
        {
            try
            {
                if (!File.Exists(settingsPath))
                {
                    return new AppSettings();
                }

                var json = File.ReadAllText(settingsPath);
                var settings = JsonSerializer.Deserialize<AppSettings>(json);
                return settings ?? new AppSettings();
            }
            catch
            {
                return new AppSettings();
            }
        }

        public void Save(AppSettings settings)
        {
            try
            {
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(settingsPath, json);
            }
            catch
            {
            }
        }
    }

    public sealed class AppSettings
    {
        public string? LastConfigIdentifier { get; set; }
    }
}
