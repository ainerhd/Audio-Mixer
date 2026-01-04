using System.Collections.Generic;

namespace Audio_Mixer.Core
{
    public sealed class ConfigLoadResult
    {
        public ConfigLoadResult(MixerSettings settings, IReadOnlyList<string> warnings)
        {
            Settings = settings;
            Warnings = warnings;
        }

        public MixerSettings Settings { get; }
        public IReadOnlyList<string> Warnings { get; }
        public bool HasWarnings => Warnings.Count > 0;
    }
}
