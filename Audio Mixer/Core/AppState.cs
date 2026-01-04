using System.Collections.Generic;

namespace Audio_Mixer.Core
{
    public sealed class AppState
    {
        public AppState(MixerSettings initialSettings)
        {
            CurrentSettings = initialSettings;
            Presets = new List<PresetDefinition>
            {
                new("Gaming", "profile:gaming"),
                new("Streaming", "profile:streaming"),
                new("Office", "profile:office"),
            };
        }

        public MixerSettings CurrentSettings { get; private set; }
        public IReadOnlyList<PresetDefinition> Presets { get; }
        public string? LastConfigIdentifier { get; private set; }

        public void ApplySettings(MixerSettings settings)
        {
            CurrentSettings = settings;
        }

        public void SetLastConfigIdentifier(string? identifier)
        {
            LastConfigIdentifier = identifier;
        }
    }

    public sealed record PresetDefinition(string Name, string Identifier);
}
