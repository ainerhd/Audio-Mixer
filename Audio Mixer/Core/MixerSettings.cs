using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.Json;

namespace Audio_Mixer.Core
{
    public sealed class MixerSettings
    {
        public int Version { get; set; } = 1;
        public int ChannelCount { get; set; } = 5;
        public int Deadzone { get; set; } = 6;
        public List<ChannelSettings> Channels { get; set; } = new();
        public bool ManualPortEnabled { get; set; }
        public string? ManualPortName { get; set; }
        public int BackgroundColorArgb { get; set; }
        public int SurfaceColorArgb { get; set; }
        public int SurfaceAccentColorArgb { get; set; }
        public int AccentColorArgb { get; set; }
        public int MutedTextColorArgb { get; set; }
        public int ChannelLabelWidth { get; set; } = 120;
        public int ChannelRowHeight { get; set; } = 52;

        public static ConfigLoadResult LoadBestEffort(string json, int maxChannels)
        {
            var defaults = CreateDefault();
            var warnings = new List<string>();

            try
            {
                using var document = JsonDocument.Parse(json);
                if (document.RootElement.ValueKind != JsonValueKind.Object)
                {
                    warnings.Add("Die Konfigurationsdatei hat kein gültiges Objektformat.");
                    return new ConfigLoadResult(defaults, warnings);
                }

                var root = document.RootElement;
                var settings = CreateDefault();

                settings.Version = ReadInt(root, "Version", 1, warnings);
                if (settings.Version > 1)
                {
                    warnings.Add($"Konfigurationsversion {settings.Version} ist neuer als erwartet.");
                }

                settings.ChannelCount = Math.Clamp(ReadInt(root, "ChannelCount", defaults.ChannelCount, warnings), 1, maxChannels);
                settings.Deadzone = Math.Clamp(ReadInt(root, "Deadzone", defaults.Deadzone, warnings), 0, 200);
                settings.ManualPortEnabled = ReadBool(root, "ManualPortEnabled", defaults.ManualPortEnabled, warnings);
                settings.ManualPortName = ReadString(root, "ManualPortName", defaults.ManualPortName, warnings);
                settings.BackgroundColorArgb = ReadInt(root, "BackgroundColorArgb", defaults.BackgroundColorArgb, warnings);
                settings.SurfaceColorArgb = ReadInt(root, "SurfaceColorArgb", defaults.SurfaceColorArgb, warnings);
                settings.SurfaceAccentColorArgb = ReadInt(root, "SurfaceAccentColorArgb", defaults.SurfaceAccentColorArgb, warnings);
                settings.AccentColorArgb = ReadInt(root, "AccentColorArgb", defaults.AccentColorArgb, warnings);
                settings.MutedTextColorArgb = ReadInt(root, "MutedTextColorArgb", defaults.MutedTextColorArgb, warnings);
                settings.ChannelLabelWidth = ReadInt(root, "ChannelLabelWidth", defaults.ChannelLabelWidth, warnings);
                settings.ChannelRowHeight = ReadInt(root, "ChannelRowHeight", defaults.ChannelRowHeight, warnings);

                settings.Channels = ReadChannels(root, settings.ChannelCount, warnings);
                return new ConfigLoadResult(settings, warnings);
            }
            catch (Exception ex)
            {
                warnings.Add($"Konfigurationsdatei konnte nicht gelesen werden: {ex.Message}");
                return new ConfigLoadResult(defaults, warnings);
            }
        }

        public static MixerSettings CreateDefault()
        {
            return new MixerSettings
            {
                Version = 1,
                ChannelCount = 5,
                Deadzone = 6,
                BackgroundColorArgb = Color.FromArgb(24, 24, 28).ToArgb(),
                SurfaceColorArgb = Color.FromArgb(36, 36, 42).ToArgb(),
                SurfaceAccentColorArgb = Color.FromArgb(44, 44, 52).ToArgb(),
                AccentColorArgb = Color.FromArgb(88, 142, 206).ToArgb(),
                MutedTextColorArgb = Color.FromArgb(180, 182, 190).ToArgb(),
                ChannelLabelWidth = 120,
                ChannelRowHeight = 52,
            };
        }

        private static List<ChannelSettings> ReadChannels(JsonElement root, int channelCount, List<string> warnings)
        {
            var channels = new List<ChannelSettings>();
            if (root.TryGetProperty("Channels", out var channelsElement)
                && channelsElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in channelsElement.EnumerateArray())
                {
                    if (item.ValueKind == JsonValueKind.Object)
                    {
                        channels.Add(new ChannelSettings
                        {
                            DeviceId = ReadString(item, "DeviceId", null, warnings),
                        });
                    }
                    else
                    {
                        channels.Add(new ChannelSettings());
                    }
                }
            }
            else
            {
                warnings.Add("Feld \"Channels\" fehlt oder ist ungültig. Standardwerte wurden verwendet.");
            }

            if (channels.Count < channelCount)
            {
                warnings.Add("Anzahl der Kanäle wurde erweitert, um der aktuellen Konfiguration zu entsprechen.");
            }

            while (channels.Count < channelCount)
            {
                channels.Add(new ChannelSettings());
            }

            if (channels.Count > channelCount)
            {
                warnings.Add("Anzahl der Kanäle wurde gekürzt, um der aktuellen Konfiguration zu entsprechen.");
                channels.RemoveRange(channelCount, channels.Count - channelCount);
            }

            return channels;
        }

        private static int ReadInt(JsonElement root, string propertyName, int fallback, List<string> warnings)
        {
            if (!root.TryGetProperty(propertyName, out var element))
            {
                warnings.Add($"Feld \"{propertyName}\" fehlt. Standardwert wurde verwendet.");
                return fallback;
            }

            return element.ValueKind switch
            {
                JsonValueKind.Number when element.TryGetInt32(out var number) => number,
                JsonValueKind.String when int.TryParse(element.GetString(), out var number) => number,
                _ => WarnAndFallback(propertyName, fallback, warnings),
            };
        }

        private static bool ReadBool(JsonElement root, string propertyName, bool fallback, List<string> warnings)
        {
            if (!root.TryGetProperty(propertyName, out var element))
            {
                warnings.Add($"Feld \"{propertyName}\" fehlt. Standardwert wurde verwendet.");
                return fallback;
            }

            return element.ValueKind switch
            {
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.String when bool.TryParse(element.GetString(), out var value) => value,
                _ => WarnAndFallback(propertyName, fallback, warnings),
            };
        }

        private static string? ReadString(JsonElement root, string propertyName, string? fallback, List<string> warnings)
        {
            if (!root.TryGetProperty(propertyName, out var element))
            {
                warnings.Add($"Feld \"{propertyName}\" fehlt. Standardwert wurde verwendet.");
                return fallback;
            }

            if (element.ValueKind == JsonValueKind.String)
            {
                return element.GetString();
            }

            warnings.Add($"Feld \"{propertyName}\" hat ein ungültiges Format. Standardwert wurde verwendet.");
            return fallback;
        }

        private static int WarnAndFallback(string propertyName, int fallback, List<string> warnings)
        {
            warnings.Add($"Feld \"{propertyName}\" hat ein ungültiges Format. Standardwert wurde verwendet.");
            return fallback;
        }

        private static bool WarnAndFallback(string propertyName, bool fallback, List<string> warnings)
        {
            warnings.Add($"Feld \"{propertyName}\" hat ein ungültiges Format. Standardwert wurde verwendet.");
            return fallback;
        }
    }

    public sealed class ChannelSettings
    {
        public string? DeviceId { get; set; }
    }
}
