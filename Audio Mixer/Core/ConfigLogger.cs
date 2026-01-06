using System;
using System.Collections.Generic;
using System.IO;

namespace Audio_Mixer.Core
{
    public static class ConfigLogger
    {
        public static void LogWarnings(string source, IReadOnlyList<string> warnings)
        {
            if (warnings.Count == 0) return;

            try
            {
                var directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Audio_Mixer");
                Directory.CreateDirectory(directory);
                var path = Path.Combine(directory, "config_load.log");

                using var writer = new StreamWriter(path, append: true);
                writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {source}");
                foreach (var warning in warnings)
                {
                    writer.WriteLine($"  - {warning}");
                }
            }
            catch
            {
            }
        }
    }
}
