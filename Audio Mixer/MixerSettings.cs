using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.Json;

namespace Audio_Mixer
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

        public static MixerSettings LoadBestEffort(string json, int maxChannels)
        {
            var defaults = CreateDefault();

            try
            {
                using var document = JsonDocument.Parse(json);
                if (document.RootElement.ValueKind != JsonValueKind.Object)
                {
                    return defaults;
                }

                var root = document.RootElement;
                var settings = CreateDefault();

                settings.Version = ReadInt(root, "Version", 1);
                settings.ChannelCount = Math.Clamp(ReadInt(root, "ChannelCount", defaults.ChannelCount), 1, maxChannels);
                settings.Deadzone = Math.Clamp(ReadInt(root, "Deadzone", defaults.Deadzone), 0, 200);
                settings.ManualPortEnabled = ReadBool(root, "ManualPortEnabled", defaults.ManualPortEnabled);
                settings.ManualPortName = ReadString(root, "ManualPortName", defaults.ManualPortName);
                settings.BackgroundColorArgb = ReadInt(root, "BackgroundColorArgb", defaults.BackgroundColorArgb);
                settings.SurfaceColorArgb = ReadInt(root, "SurfaceColorArgb", defaults.SurfaceColorArgb);
                settings.SurfaceAccentColorArgb = ReadInt(root, "SurfaceAccentColorArgb", defaults.SurfaceAccentColorArgb);
                settings.AccentColorArgb = ReadInt(root, "AccentColorArgb", defaults.AccentColorArgb);
                settings.MutedTextColorArgb = ReadInt(root, "MutedTextColorArgb", defaults.MutedTextColorArgb);
                settings.ChannelLabelWidth = ReadInt(root, "ChannelLabelWidth", defaults.ChannelLabelWidth);
                settings.ChannelRowHeight = ReadInt(root, "ChannelRowHeight", defaults.ChannelRowHeight);

                settings.Channels = ReadChannels(root, settings.ChannelCount);
                return settings;
            }
            catch
            {
                return defaults;
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

        private static List<ChannelSettings> ReadChannels(JsonElement root, int channelCount)
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
                            DeviceId = ReadString(item, "DeviceId", null),
                        });
                    }
                    else
                    {
                        channels.Add(new ChannelSettings());
                    }
                }
            }

            while (channels.Count < channelCount)
            {
                channels.Add(new ChannelSettings());
            }

            if (channels.Count > channelCount)
            {
                channels.RemoveRange(channelCount, channels.Count - channelCount);
            }

            return channels;
        }

        private static int ReadInt(JsonElement root, string propertyName, int fallback)
        {
            if (!root.TryGetProperty(propertyName, out var element))
            {
                return fallback;
            }

            return element.ValueKind switch
            {
                JsonValueKind.Number when element.TryGetInt32(out var number) => number,
                JsonValueKind.String when int.TryParse(element.GetString(), out var number) => number,
                _ => fallback,
            };
        }

        private static bool ReadBool(JsonElement root, string propertyName, bool fallback)
        {
            if (!root.TryGetProperty(propertyName, out var element))
            {
                return fallback;
            }

            return element.ValueKind switch
            {
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.String when bool.TryParse(element.GetString(), out var value) => value,
                _ => fallback,
            };
        }

        private static string? ReadString(JsonElement root, string propertyName, string? fallback)
        {
            if (!root.TryGetProperty(propertyName, out var element))
            {
                return fallback;
            }

            return element.ValueKind == JsonValueKind.String ? element.GetString() : fallback;
        }
    }

    public sealed class ChannelSettings
    {
        public string? DeviceId { get; set; }
    }
}
