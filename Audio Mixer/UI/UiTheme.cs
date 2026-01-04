using System.Drawing;
using Audio_Mixer.Core;

namespace Audio_Mixer.UI
{
    public sealed class UiTheme
    {
        public UiTheme(
            Color background,
            Color surface,
            Color surfaceAlt,
            Color accent,
            Color text,
            Color mutedText,
            Color warningBackground,
            Color warningText)
        {
            Background = background;
            Surface = surface;
            SurfaceAlt = surfaceAlt;
            Accent = accent;
            Text = text;
            MutedText = mutedText;
            WarningBackground = warningBackground;
            WarningText = warningText;
        }

        public Color Background { get; }
        public Color Surface { get; }
        public Color SurfaceAlt { get; }
        public Color Accent { get; }
        public Color Text { get; }
        public Color MutedText { get; }
        public Color WarningBackground { get; }
        public Color WarningText { get; }

        public Font BaseFont { get; } = new("Segoe UI", 9.5f, FontStyle.Regular);
        public Font HeaderFont { get; } = new("Segoe UI Semibold", 14f, FontStyle.Bold);
        public Font SectionFont { get; } = new("Segoe UI Semibold", 11f, FontStyle.Bold);
        public Font LabelFont { get; } = new("Segoe UI Semibold", 9.5f, FontStyle.Bold);

        public Padding PagePadding { get; } = new(20);
        public Padding CardPadding { get; } = new(16);
        public Padding CompactPadding { get; } = new(8, 6, 8, 6);

        public static UiTheme FromSettings(MixerSettings settings)
        {
            return new UiTheme(
                Color.FromArgb(settings.BackgroundColorArgb),
                Color.FromArgb(settings.SurfaceColorArgb),
                Color.FromArgb(settings.SurfaceAccentColorArgb),
                Color.FromArgb(settings.AccentColorArgb),
                Color.White,
                Color.FromArgb(settings.MutedTextColorArgb),
                Color.FromArgb(80, 68, 24),
                Color.FromArgb(255, 231, 160));
        }
    }
}
