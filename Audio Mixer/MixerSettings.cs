using System.Collections.Generic;
using System.Drawing;

namespace Audio_Mixer
{
    public sealed class MixerSettings
    {
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

        public static MixerSettings CreateDefault()
        {
            return new MixerSettings
            {
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
    }

    public sealed class ChannelSettings
    {
        public string? DeviceId { get; set; }
    }
}
