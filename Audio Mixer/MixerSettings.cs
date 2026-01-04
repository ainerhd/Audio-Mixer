using System.Collections.Generic;

namespace Audio_Mixer
{
    public sealed class MixerSettings
    {
        public int ChannelCount { get; set; } = 5;
        public int Deadzone { get; set; } = 6;
        public List<ChannelSettings> Channels { get; set; } = new();

        public static MixerSettings CreateDefault()
        {
            return new MixerSettings
            {
                ChannelCount = 5,
                Deadzone = 6,
            };
        }
    }

    public sealed class ChannelSettings
    {
        public string? DeviceId { get; set; }
    }
}
