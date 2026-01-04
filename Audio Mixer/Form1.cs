using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;

namespace Audio_Mixer
{
    public partial class Form1 : Form
    {
        private const int MaxChannels = 8;

        private readonly CoreAudioManager audioManager = new();
        private readonly List<ChannelRow> channelRows = new();
        private readonly List<AudioDeviceItem> audioDevices = new();

        private CancellationTokenSource? scanCts;
        private SerialPort? serialPort;
        private int[] lastValues = Array.Empty<int>();
        private MixerSettings settings = MixerSettings.CreateDefault();
        private bool isApplyingSettings;

        private Panel topBar = null!;
        private Button mixerNavButton = null!;
        private Button settingsNavButton = null!;
        private Panel mixerNavUnderline = null!;
        private Panel settingsNavUnderline = null!;
        private Panel contentPanel = null!;
        private Panel mixerViewPanel = null!;
        private Panel settingsViewPanel = null!;
        private Label statusLabel = null!;
        private Button rescanButton = null!;
        private NumericUpDown channelCountUpDown = null!;
        private NumericUpDown deadzoneUpDown = null!;
        private TableLayoutPanel channelsTable = null!;
        private Button saveSettingsButton = null!;
        private Button loadSettingsButton = null!;
        private CheckBox manualPortCheckBox = null!;
        private ComboBox manualPortComboBox = null!;
        private Button manualConnectButton = null!;
        private Button refreshPortsButton = null!;
        private NumericUpDown channelRowHeightUpDown = null!;
        private NumericUpDown channelLabelWidthUpDown = null!;
        private TableLayoutPanel channelsHeaderTable = null!;
        private Panel channelsScrollPanel = null!;
        private StatusDot statusDot = null!;
        private ContextMenuStrip profileMenu = null!;
        private Label manualPortHintLabel = null!;

        private Color ThemeBack => Color.FromArgb(settings.BackgroundColorArgb);
        private Color ThemeSurface => Color.FromArgb(settings.SurfaceColorArgb);
        private Color ThemeSurface2 => Color.FromArgb(settings.SurfaceAccentColorArgb);
        private Color ThemeAccent => Color.FromArgb(settings.AccentColorArgb);
        private Color ThemeText => Color.White;
        private Color ThemeMutedText => Color.FromArgb(settings.MutedTextColorArgb);

        private readonly ToolTip deviceToolTip = new() { ShowAlways = true };
        private StatusState statusState = StatusState.Idle;
        private ViewKind currentView = ViewKind.Mixer;

        public Form1()
        {
            InitializeComponent();
            BuildUi();
            UpdateDevices();
            if (!TryAutoLoadLastConfig())
            {
                ApplySettings(settings);
            }
            StartAutoScan();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            scanCts?.Cancel();
            CloseSerialPort();
            audioManager.Dispose();
        }

        private void BuildUi()
        {
            Text = "Audio Mixer";
            MinimumSize = new Size(720, 480);
            BackColor = ThemeBack;
            ForeColor = ThemeText;
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
            DoubleBuffered = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            UpdateStyles();

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = ThemeBack,
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            Controls.Add(root);

            BuildTopBar(root);

            contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ThemeBack,
            };
            root.Controls.Add(contentPanel, 0, 1);

            mixerViewPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ThemeBack,
            };
            settingsViewPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ThemeBack,
                Visible = false,
            };
            contentPanel.Controls.Add(mixerViewPanel);
            contentPanel.Controls.Add(settingsViewPanel);

            BuildMixerTabContent(mixerViewPanel);
            BuildSettingsTabContent(settingsViewPanel);

            SetActiveView(currentView);
        }

        private void BuildTopBar(TableLayoutPanel root)
        {
            topBar = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ThemeSurface,
                Padding = new Padding(20, 10, 20, 8),
            };
            root.Controls.Add(topBar, 0, 0);

            var navLayout = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0),
            };
            topBar.Controls.Add(navLayout);

            navLayout.Controls.Add(CreateNavItem("Mixer", out mixerNavButton, out mixerNavUnderline));
            navLayout.Controls.Add(CreateNavItem("Einstellungen", out settingsNavButton, out settingsNavUnderline));

            mixerNavButton.Click += (_, _) => SetActiveView(ViewKind.Mixer);
            settingsNavButton.Click += (_, _) => SetActiveView(ViewKind.Settings);
        }

        private Control CreateNavItem(string text, out Button button, out Panel underline)
        {
            var container = new TableLayoutPanel
            {
                AutoSize = true,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0, 0, 12, 0),
            };
            container.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            container.RowStyles.Add(new RowStyle(SizeType.Absolute, 2));

            button = new Button
            {
                Text = text,
                AutoSize = true,
                BackColor = ThemeSurface,
                FlatStyle = FlatStyle.Flat,
                ForeColor = ThemeMutedText,
                Padding = new Padding(12, 6, 12, 6),
                Margin = new Padding(0),
            };
            button.FlatAppearance.BorderSize = 0;
            container.Controls.Add(button, 0, 0);

            underline = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ThemeAccent,
                Visible = false,
                Margin = new Padding(0),
            };
            container.Controls.Add(underline, 0, 1);

            return container;
        }

        private void SetActiveView(ViewKind view)
        {
            currentView = view;
            var isMixer = view == ViewKind.Mixer;
            mixerViewPanel.Visible = isMixer;
            settingsViewPanel.Visible = !isMixer;

            UpdateNavState(mixerNavButton, mixerNavUnderline, isMixer);
            UpdateNavState(settingsNavButton, settingsNavUnderline, !isMixer);
        }

        private void UpdateNavState(Button button, Panel underline, bool isActive)
        {
            button.BackColor = isActive ? ThemeSurface2 : ThemeSurface;
            button.ForeColor = isActive ? ThemeText : ThemeMutedText;
            underline.Visible = isActive;
        }

        private void BuildMixerTabContent(Panel tab)
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(20),
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            tab.Controls.Add(root);

            BuildHeader(root);
            BuildChannelContainer(root);
        }

        private void BuildHeader(TableLayoutPanel root)
        {
            var headerCard = CreateCardPanel();
            headerCard.Margin = new Padding(0, 0, 0, 16);
            root.Controls.Add(headerCard);

            var headerPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                AutoSize = true,
            };
            headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            headerCard.Controls.Add(headerPanel);

            var titleLabel = new Label
            {
                Text = "Mixer Verbindung",
                Font = new Font("Segoe UI Semibold", 14f, FontStyle.Bold),
                AutoSize = true,
                Dock = DockStyle.Fill,
                ForeColor = ThemeText,
            };
            headerPanel.Controls.Add(titleLabel, 0, 0);

            var statusPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0),
            };
            statusDot = new StatusDot
            {
                Size = new Size(12, 12),
                Margin = new Padding(0, 4, 6, 0),
                DotColor = GetStatusColor(statusState),
            };
            statusLabel = new Label
            {
                Text = "Status: Nicht verbunden",
                AutoSize = true,
                ForeColor = ThemeMutedText,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
            };
            statusPanel.Controls.Add(statusDot);
            statusPanel.Controls.Add(statusLabel);
            headerPanel.Controls.Add(statusPanel, 1, 0);

            var actionPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0),
                Padding = new Padding(0),
            };

            rescanButton = CreateHeaderButton("Auto-Suche");
            rescanButton.Click += (_, _) => StartAutoScan();
            actionPanel.Controls.Add(rescanButton);

            headerPanel.Controls.Add(actionPanel, 2, 0);
        }

        private void BuildChannelContainer(TableLayoutPanel root)
        {
            var channelsCard = CreateCardPanel();
            root.Controls.Add(channelsCard);

            var channelsContainer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
            };
            channelsContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            channelsContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            channelsContainer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            channelsCard.Controls.Add(channelsContainer);

            var channelsHeader = new Label
            {
                Text = "Kanäle & Pegel",
                AutoSize = true,
                Font = new Font("Segoe UI Semibold", 11f, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 8),
                ForeColor = ThemeText,
            };
            channelsContainer.Controls.Add(channelsHeader, 0, 0);

            channelsHeaderTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 4,
                AutoSize = true,
            };
            ApplyChannelColumnStyles(channelsHeaderTable);
            channelsHeaderTable.Controls.Add(CreateHeaderLabel("Kanal"), 0, 0);
            channelsHeaderTable.Controls.Add(CreateHeaderLabel("Audio-Ausgang"), 1, 0);
            channelsHeaderTable.Controls.Add(CreateHeaderLabel("Level"), 2, 0);
            channelsHeaderTable.Controls.Add(CreateHeaderLabel("Mute"), 3, 0);
            channelsContainer.Controls.Add(channelsHeaderTable, 0, 1);

            channelsScrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(0, 4, 0, 0),
            };
            channelsContainer.Controls.Add(channelsScrollPanel, 0, 2);

            channelsTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 4,
                AutoSize = true,
            };
            ApplyChannelColumnStyles(channelsTable);
            channelsScrollPanel.Controls.Add(channelsTable);
        }

        private void BuildSettingsTabContent(Panel tab)
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(20),
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tab.Controls.Add(root);

            var generalCard = CreateCardPanel();
            generalCard.Margin = new Padding(0, 0, 0, 16);
            root.Controls.Add(generalCard);

            var generalContainer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                AutoSize = true,
            };
            generalContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            generalContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            generalContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            generalCard.Controls.Add(generalContainer);

            generalContainer.Controls.Add(CreateSectionHeader("Allgemeine Einstellungen"), 0, 0);

            var settingsPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 5,
                AutoSize = true,
            };
            settingsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            settingsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            settingsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            settingsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            settingsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            generalContainer.Controls.Add(settingsPanel, 0, 1);

            settingsPanel.Controls.Add(CreateFieldLabel("Kanäle:"), 0, 0);
            channelCountUpDown = new NumericUpDown
            {
                Minimum = 1,
                Maximum = MaxChannels,
                Value = settings.ChannelCount,
                Dock = DockStyle.Fill,
                BackColor = ThemeSurface2,
                ForeColor = ThemeText,
            };
            channelCountUpDown.ValueChanged += (_, _) =>
            {
                if (isApplyingSettings) return;
                settings.ChannelCount = (int)channelCountUpDown.Value;
                BuildChannelRows(settings.ChannelCount);
            };
            settingsPanel.Controls.Add(channelCountUpDown, 1, 0);

            settingsPanel.Controls.Add(CreateFieldLabel("Deadzone:"), 2, 0);
            deadzoneUpDown = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 200,
                Value = settings.Deadzone,
                Dock = DockStyle.Fill,
                BackColor = ThemeSurface2,
                ForeColor = ThemeText,
            };
            deadzoneUpDown.ValueChanged += (_, _) =>
            {
                if (isApplyingSettings) return;
                settings.Deadzone = (int)deadzoneUpDown.Value;
            };
            settingsPanel.Controls.Add(deadzoneUpDown, 3, 0);

            var configPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ThemeSurface2,
                Padding = new Padding(8, 6, 8, 6),
                Margin = new Padding(8, 0, 0, 0),
            };
            var configLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
            };
            configLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            configLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            configPanel.Controls.Add(configLayout);

            loadSettingsButton = new Button
            {
                Text = "Laden",
                Dock = DockStyle.Fill,
                BackColor = ThemeSurface,
                FlatStyle = FlatStyle.Flat,
                ForeColor = ThemeText,
                Margin = new Padding(0, 0, 6, 0),
                Padding = new Padding(6, 2, 6, 2),
            };
            loadSettingsButton.FlatAppearance.BorderSize = 0;
            loadSettingsButton.Click += (_, _) => LoadSettingsFromFile();
            configLayout.Controls.Add(loadSettingsButton, 0, 0);

            saveSettingsButton = new Button
            {
                Text = "Speichern",
                Dock = DockStyle.Fill,
                BackColor = ThemeSurface,
                FlatStyle = FlatStyle.Flat,
                ForeColor = ThemeText,
                Margin = new Padding(6, 0, 0, 0),
                Padding = new Padding(6, 2, 6, 2),
            };
            saveSettingsButton.FlatAppearance.BorderSize = 0;
            saveSettingsButton.Click += (_, _) => SaveSettingsToFile();
            configLayout.Controls.Add(saveSettingsButton, 1, 0);

            settingsPanel.Controls.Add(configPanel, 4, 0);

            profileMenu = new ContextMenuStrip();
            profileMenu.Items.Add("Gaming", null, (_, _) => LoadProfile("gaming", true));
            profileMenu.Items.Add("Streaming", null, (_, _) => LoadProfile("streaming", true));
            profileMenu.Items.Add("Office", null, (_, _) => LoadProfile("office", true));
            profileMenu.Items.Add(new ToolStripSeparator());
            profileMenu.Items.Add("Aus Datei laden...", null, (_, _) => LoadSettingsFromFile());
            profileMenu.Items.Add("Aktuelle Einstellungen speichern...", null, (_, _) => SaveSettingsToFile());

            var profileButton = new Button
            {
                Text = "Profile ▾",
                AutoSize = true,
                BackColor = ThemeAccent,
                FlatStyle = FlatStyle.Flat,
                ForeColor = ThemeText,
                Margin = new Padding(0, 8, 0, 0),
                Padding = new Padding(8, 4, 8, 4),
            };
            profileButton.FlatAppearance.BorderSize = 0;
            profileButton.Click += (_, _) => profileMenu.Show(profileButton, new Point(0, profileButton.Height));

            var profilePanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0),
            };
            profilePanel.Controls.Add(profileButton);
            generalContainer.Controls.Add(profilePanel, 0, 2);

            var connectionCard = CreateCardPanel();
            connectionCard.Margin = new Padding(0, 0, 0, 16);
            root.Controls.Add(connectionCard);

            var connectionContainer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                AutoSize = true,
            };
            connectionContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            connectionContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            connectionContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            connectionCard.Controls.Add(connectionContainer);

            connectionContainer.Controls.Add(CreateSectionHeader("Verbindung"), 0, 0);

            var connectionPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                AutoSize = true,
            };
            connectionPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            connectionPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            connectionPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            connectionPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            connectionContainer.Controls.Add(connectionPanel, 0, 1);

            manualPortCheckBox = new CheckBox
            {
                Text = "Manuelle Portwahl aktivieren",
                Dock = DockStyle.Fill,
                ForeColor = ThemeText,
                AutoSize = true,
            };
            manualPortCheckBox.CheckedChanged += (_, _) =>
            {
                if (isApplyingSettings) return;
                settings.ManualPortEnabled = manualPortCheckBox.Checked;
                UpdateManualPortUi();
            };
            connectionPanel.Controls.Add(manualPortCheckBox, 0, 0);

            manualPortComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = ThemeSurface2,
                ForeColor = ThemeText,
                FlatStyle = FlatStyle.Flat,
            };
            manualPortComboBox.SelectedIndexChanged += (_, _) =>
            {
                if (manualPortComboBox.SelectedItem is string portName)
                {
                    settings.ManualPortName = portName;
                }
            };
            connectionPanel.Controls.Add(manualPortComboBox, 1, 0);

            refreshPortsButton = new Button
            {
                Text = "Ports aktualisieren",
                Dock = DockStyle.Fill,
                BackColor = ThemeSurface2,
                FlatStyle = FlatStyle.Flat,
                ForeColor = ThemeText,
            };
            refreshPortsButton.FlatAppearance.BorderSize = 0;
            refreshPortsButton.Click += (_, _) => PopulatePortList();
            connectionPanel.Controls.Add(refreshPortsButton, 2, 0);

            manualConnectButton = new Button
            {
                Text = "Verbinden",
                Dock = DockStyle.Fill,
                BackColor = ThemeAccent,
                FlatStyle = FlatStyle.Flat,
                ForeColor = ThemeText,
            };
            manualConnectButton.FlatAppearance.BorderSize = 0;
            manualConnectButton.Click += (_, _) => ConnectManualPort();
            connectionPanel.Controls.Add(manualConnectButton, 3, 0);

            manualPortHintLabel = new Label
            {
                Text = "Aktiviere manuelle Portwahl, um einen Port auszuwählen.",
                AutoSize = true,
                ForeColor = ThemeMutedText,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 6, 0, 0),
            };
            connectionContainer.Controls.Add(manualPortHintLabel, 0, 2);

            var colorCard = CreateCardPanel();
            colorCard.Margin = new Padding(0, 0, 0, 16);
            root.Controls.Add(colorCard);

            var colorContainer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                AutoSize = true,
            };
            colorContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            colorContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            colorCard.Controls.Add(colorContainer);

            colorContainer.Controls.Add(CreateSectionHeader("Farben"), 0, 0);

            var colorPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                AutoSize = true,
            };
            colorPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            colorPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            colorPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));
            colorContainer.Controls.Add(colorPanel, 0, 1);

            colorPanel.Controls.Add(CreateHeaderLabel("Farbe"), 0, 0);
            colorPanel.Controls.Add(CreateHeaderLabel("Vorschau"), 1, 0);
            colorPanel.Controls.Add(CreateHeaderLabel("Aktion"), 2, 0);

            AddColorPickerRow(colorPanel, 1, "Hintergrund", () => ThemeBack,
                color => settings.BackgroundColorArgb = color.ToArgb());
            AddColorPickerRow(colorPanel, 2, "Kartenfläche", () => ThemeSurface,
                color => settings.SurfaceColorArgb = color.ToArgb());
            AddColorPickerRow(colorPanel, 3, "Kartenakzent", () => ThemeSurface2,
                color => settings.SurfaceAccentColorArgb = color.ToArgb());
            AddColorPickerRow(colorPanel, 4, "Akzentfarbe", () => ThemeAccent,
                color => settings.AccentColorArgb = color.ToArgb());
            AddColorPickerRow(colorPanel, 5, "Gedämpfter Text", () => ThemeMutedText,
                color => settings.MutedTextColorArgb = color.ToArgb());

            var channelSizeCard = CreateCardPanel();
            root.Controls.Add(channelSizeCard);

            var channelSizeContainer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                AutoSize = true,
            };
            channelSizeContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            channelSizeContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            channelSizeCard.Controls.Add(channelSizeContainer);

            channelSizeContainer.Controls.Add(CreateSectionHeader("Kanalgröße"), 0, 0);

            var channelSizePanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                AutoSize = true,
            };
            channelSizePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            channelSizePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            channelSizePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            channelSizePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            channelSizeContainer.Controls.Add(channelSizePanel, 0, 1);

            channelSizePanel.Controls.Add(CreateFieldLabel("Zeilenhöhe:"), 0, 0);
            channelRowHeightUpDown = new NumericUpDown
            {
                Minimum = 40,
                Maximum = 120,
                Value = settings.ChannelRowHeight,
                Dock = DockStyle.Fill,
                BackColor = ThemeSurface2,
                ForeColor = ThemeText,
            };
            channelRowHeightUpDown.ValueChanged += (_, _) =>
            {
                if (isApplyingSettings) return;
                settings.ChannelRowHeight = (int)channelRowHeightUpDown.Value;
                BuildChannelRows(settings.ChannelCount);
            };
            channelSizePanel.Controls.Add(channelRowHeightUpDown, 1, 0);

            channelSizePanel.Controls.Add(CreateFieldLabel("Kanalbreite:"), 2, 0);
            channelLabelWidthUpDown = new NumericUpDown
            {
                Minimum = 80,
                Maximum = 240,
                Value = settings.ChannelLabelWidth,
                Dock = DockStyle.Fill,
                BackColor = ThemeSurface2,
                ForeColor = ThemeText,
            };
            channelLabelWidthUpDown.ValueChanged += (_, _) =>
            {
                if (isApplyingSettings) return;
                settings.ChannelLabelWidth = (int)channelLabelWidthUpDown.Value;
                BuildChannelRows(settings.ChannelCount);
            };
            channelSizePanel.Controls.Add(channelLabelWidthUpDown, 3, 0);
        }

        private Label CreateFieldLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                ForeColor = ThemeMutedText,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = true,
            };
        }

        private Label CreateSectionHeader(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI Semibold", 11f, FontStyle.Bold),
                ForeColor = ThemeText,
                AutoSize = true,
            };
        }

        private Panel CreateCardPanel()
        {
            return new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ThemeSurface,
                Padding = new Padding(16),
            };
        }

        private Button CreateHeaderButton(string text)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = true,
                BackColor = ThemeAccent,
                FlatStyle = FlatStyle.Flat,
                ForeColor = ThemeText,
                Margin = new Padding(4, 0, 0, 0),
                Padding = new Padding(8, 4, 8, 4),
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private void ApplyChannelColumnStyles(TableLayoutPanel table)
        {
            table.ColumnStyles.Clear();
            table.ColumnCount = 4;
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, settings.ChannelLabelWidth));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 52));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        }

        private static Color GetStatusColor(StatusState state)
        {
            return state switch
            {
                StatusState.Searching => Color.Goldenrod,
                StatusState.Connected => Color.LimeGreen,
                StatusState.Error => Color.IndianRed,
                _ => Color.DimGray,
            };
        }

        private void UpdateStatus(StatusState state, string text)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired)
            {
                BeginInvoke(() => UpdateStatus(state, text));
                return;
            }

            statusState = state;
            if (statusLabel != null)
            {
                statusLabel.Text = text;
            }

            if (statusDot != null)
            {
                statusDot.DotColor = GetStatusColor(state);
                statusDot.Invalidate();
            }
        }

        private void UpdateDeviceToolTip(ComboBox comboBox)
        {
            if (comboBox.SelectedItem is AudioDeviceItem item)
            {
                deviceToolTip.SetToolTip(comboBox, item.Name);
            }
            else
            {
                deviceToolTip.SetToolTip(comboBox, string.Empty);
            }
        }

        private void UpdateComboDropDownWidth(ComboBox comboBox)
        {
            if (comboBox.Items.Count == 0) return;

            var maxWidth = comboBox.Width;
            foreach (var item in comboBox.Items)
            {
                var text = item?.ToString() ?? string.Empty;
                var size = TextRenderer.MeasureText(text, comboBox.Font);
                maxWidth = Math.Max(maxWidth, size.Width);
            }

            maxWidth += SystemInformation.VerticalScrollBarWidth + 16;
            comboBox.DropDownWidth = maxWidth;
        }

        private void ToggleMute(ChannelRow row)
        {
            row.IsMuted = !row.IsMuted;
            UpdateMuteUi(row);

            var deviceId = settings.Channels.ElementAtOrDefault(row.Index)?.DeviceId;
            if (string.IsNullOrWhiteSpace(deviceId)) return;

            if (row.IsMuted)
            {
                audioManager.SetDeviceVolume(deviceId, 0f);
                return;
            }

            var value = row.Index < lastValues.Length ? lastValues[row.Index] : 0;
            if (value < 0) value = 0;
            var volume = value / 1023f;
            audioManager.SetDeviceVolume(deviceId, volume);
        }

        private void UpdateMuteUi(ChannelRow row)
        {
            if (row.IsMuted)
            {
                row.MuteButton.Text = "Muted";
                row.MuteButton.BackColor = Color.FromArgb(96, 38, 38);
                row.MuteButton.ForeColor = ThemeText;
                row.LevelPercentLabel.ForeColor = ThemeMutedText;
                row.LevelFill.BackColor = Color.FromArgb(72, 72, 72);
            }
            else
            {
                row.MuteButton.Text = "Mute";
                row.MuteButton.BackColor = ThemeSurface2;
                row.MuteButton.ForeColor = ThemeText;
                row.LevelPercentLabel.ForeColor = ThemeMutedText;
                row.LevelFill.BackColor = ThemeAccent;
            }
        }

        private void LoadProfile(string profileName, bool persist)
        {
            var profilePath = Path.Combine(AppContext.BaseDirectory, $"{profileName}.json");
            MixerSettings? profile = null;

            if (TryReadSettingsFromPath(profilePath, false, out var loaded))
            {
                profile = loaded;
            }

            profile ??= GetFallbackProfile(profileName);
            settings = profile;
            RebuildUi();

            if (persist)
            {
                PersistLastConfigIdentifier($"profile:{profileName}");
            }
        }

        private static MixerSettings GetFallbackProfile(string profileName)
        {
            var fallback = MixerSettings.CreateDefault();
            return profileName.ToLowerInvariant() switch
            {
                "gaming" => fallback,
                "streaming" => fallback,
                "office" => fallback,
                _ => fallback,
            };
        }

        private static Panel CreateCellPanel(Control content, Color backColor)
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = backColor,
                Padding = new Padding(8, 6, 8, 6),
            };
            content.Dock = DockStyle.Fill;
            panel.Controls.Add(content);
            return panel;
        }

        private Label CreateHeaderLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI Semibold", 9.5f, FontStyle.Bold),
                ForeColor = ThemeMutedText,
                TextAlign = ContentAlignment.MiddleLeft,
            };
        }

        private void AddColorPickerRow(TableLayoutPanel panel, int rowIndex, string labelText, Func<Color> getColor, Action<Color> setColor)
        {
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var label = CreateFieldLabel(labelText);
            var preview = new Panel
            {
                BackColor = getColor(),
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 6, 0, 6),
                BorderStyle = BorderStyle.FixedSingle,
            };
            var button = new Button
            {
                Text = "Farbe wählen",
                Dock = DockStyle.Fill,
                BackColor = ThemeSurface2,
                FlatStyle = FlatStyle.Flat,
                ForeColor = ThemeText,
            };
            button.FlatAppearance.BorderSize = 0;
            button.Click += (_, _) =>
            {
                using var dialog = new ColorDialog
                {
                    Color = getColor(),
                    FullOpen = true,
                };
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                setColor(dialog.Color);
                RebuildUi();
            };

            panel.Controls.Add(label, 0, rowIndex);
            panel.Controls.Add(preview, 1, rowIndex);
            panel.Controls.Add(button, 2, rowIndex);
        }

        private void RebuildUi()
        {
            var selectedView = currentView;
            var statusText = statusLabel?.Text;
            var previousStatus = statusState;

            SuspendLayout();
            Controls.Clear();
            BuildUi();
            UpdateDevices();
            ApplySettings(settings);

            if (!string.IsNullOrEmpty(statusText))
            {
                UpdateStatus(previousStatus, statusText);
            }

            SetActiveView(selectedView);

            ResumeLayout();
        }

        private void PopulatePortList()
        {
            if (manualPortComboBox == null) return;

            manualPortComboBox.Items.Clear();
            foreach (var port in SerialPort.GetPortNames().OrderBy(name => name))
            {
                manualPortComboBox.Items.Add(port);
            }

            if (!string.IsNullOrWhiteSpace(settings.ManualPortName)
                && manualPortComboBox.Items.Contains(settings.ManualPortName))
            {
                manualPortComboBox.SelectedItem = settings.ManualPortName;
            }
            else if (manualPortComboBox.Items.Count > 0)
            {
                manualPortComboBox.SelectedIndex = 0;
            }
        }

        private void UpdateManualPortUi()
        {
            var manualEnabled = settings.ManualPortEnabled;
            manualPortComboBox.Enabled = manualEnabled;
            manualConnectButton.Enabled = manualEnabled;
            refreshPortsButton.Enabled = manualEnabled;
            rescanButton.Enabled = !manualEnabled;
            manualPortHintLabel.Visible = !manualEnabled;

            if (manualEnabled)
            {
                scanCts?.Cancel();
                UpdateStatus(StatusState.Idle, "Status: Manuelle Portwahl aktiv");
            }
        }

        private void ConnectManualPort()
        {
            if (!settings.ManualPortEnabled)
            {
                UpdateStatus(StatusState.Idle, "Status: Manuelle Portwahl deaktiviert");
                return;
            }

            var portName = settings.ManualPortName;
            if (string.IsNullOrWhiteSpace(portName))
            {
                UpdateStatus(StatusState.Idle, "Status: Kein Port ausgewählt");
                return;
            }

            ConnectToPort(portName);
        }

        private void BuildChannelRows(int count)
        {
            channelsTable.SuspendLayout();
            channelsTable.Controls.Clear();
            channelsTable.RowStyles.Clear();
            channelRows.Clear();

            lastValues = Enumerable.Repeat(-1, count).ToArray();

            ApplyChannelColumnStyles(channelsTable);
            ApplyChannelColumnStyles(channelsHeaderTable);

            if (settings.Channels.Count > count)
            {
                settings.Channels.RemoveRange(count, settings.Channels.Count - count);
            }

            for (var i = 0; i < count; i++)
            {
                var rowIndex = i;
                var rowBackground = i % 2 == 0 ? ThemeSurface2 : ThemeSurface;
                channelsTable.RowStyles.Add(new RowStyle(SizeType.Absolute, settings.ChannelRowHeight));

                var label = new Label
                {
                    Text = $"Kanal {i + 1}",
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleLeft,
                    ForeColor = ThemeText,
                };
                channelsTable.Controls.Add(CreateCellPanel(label, rowBackground), 0, rowIndex);

                var deviceCombo = new ComboBox
                {
                    Dock = DockStyle.Fill,
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    BackColor = ThemeSurface,
                    ForeColor = ThemeText,
                    FlatStyle = FlatStyle.Flat,
                };

                var indexCopy = i; // wichtig: Closure-Falle vermeiden
                deviceCombo.SelectedIndexChanged += (_, _) =>
                {
                    if (deviceCombo.SelectedItem is AudioDeviceItem item)
                    {
                        EnsureChannelSettings(indexCopy);
                        settings.Channels[indexCopy].DeviceId = item.Id;
                    }
                    UpdateDeviceToolTip(deviceCombo);
                };
                deviceCombo.SizeChanged += (_, _) => UpdateComboDropDownWidth(deviceCombo);

                channelsTable.Controls.Add(CreateCellPanel(deviceCombo, rowBackground), 1, rowIndex);

                var progressTrack = new Panel
                {
                    Dock = DockStyle.Fill,
                    BackColor = ThemeSurface2,
                    Padding = new Padding(2),
                    Margin = new Padding(0),
                };
                var progressFill = new Panel
                {
                    Dock = DockStyle.Left,
                    BackColor = ThemeAccent,
                    Width = 0,
                };
                progressTrack.Controls.Add(progressFill);

                var levelPercentLabel = new Label
                {
                    Text = "0%",
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleRight,
                    ForeColor = ThemeMutedText,
                    AutoSize = false,
                };
                var muteButton = new Button
                {
                    Text = "Mute",
                    Dock = DockStyle.Fill,
                    MinimumSize = new Size(72, 24),
                    BackColor = ThemeSurface2,
                    FlatStyle = FlatStyle.Flat,
                    ForeColor = ThemeText,
                    Margin = new Padding(6, 6, 6, 6),
                    Padding = new Padding(4, 2, 4, 2),
                };
                muteButton.FlatAppearance.BorderSize = 0;

                var progressPanel = new Panel
                {
                    Dock = DockStyle.Fill,
                    Padding = new Padding(0, Math.Max(6, (settings.ChannelRowHeight - 22) / 2), 0,
                        Math.Max(6, (settings.ChannelRowHeight - 22) / 2)),
                };
                progressPanel.Controls.Add(progressTrack);

                var levelContainer = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 2,
                    RowCount = 1,
                };
                levelContainer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                levelContainer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 52));
                levelContainer.Controls.Add(progressPanel, 0, 0);
                levelContainer.Controls.Add(levelPercentLabel, 1, 0);
                channelsTable.Controls.Add(CreateCellPanel(levelContainer, rowBackground), 2, rowIndex);
                channelsTable.Controls.Add(CreateCellPanel(muteButton, rowBackground), 3, rowIndex);

                var row = new ChannelRow(i, deviceCombo, progressTrack, progressFill, levelPercentLabel, muteButton);
                muteButton.Click += (_, _) => ToggleMute(row);
                progressTrack.SizeChanged += (_, _) => UpdateLevelBar(row, row.LevelValue);
                channelRows.Add(row);
                UpdateMuteUi(row);
            }

            channelsTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            channelsTable.RowCount = count + 1;
            var fillerPanel = new Panel { Dock = DockStyle.Fill, Visible = false };
            channelsTable.Controls.Add(fillerPanel, 0, count);
            channelsTable.SetColumnSpan(fillerPanel, channelsTable.ColumnCount);

            PopulateDeviceCombos();
            ApplyChannelSelections();
            channelsTable.ResumeLayout();
        }

        private void UpdateDevices()
        {
            audioDevices.Clear();
            foreach (var device in audioManager.GetOutputDevices())
            {
                audioDevices.Add(new AudioDeviceItem(device.Id, device.Name));
            }
            PopulateDeviceCombos();
        }

        private void PopulateDeviceCombos()
        {
            foreach (var row in channelRows)
            {
                row.DeviceComboBox.Items.Clear();
                foreach (var device in audioDevices)
                {
                    row.DeviceComboBox.Items.Add(device);
                }
                UpdateComboDropDownWidth(row.DeviceComboBox);
            }
        }

        private void ApplySettings(MixerSettings newSettings)
        {
            settings = newSettings;

            if (settings.ChannelCount < 1) settings.ChannelCount = 1;
            if (settings.ChannelCount > MaxChannels) settings.ChannelCount = MaxChannels;
            if (settings.Deadzone < 0) settings.Deadzone = 0;
            if (settings.Deadzone > deadzoneUpDown.Maximum) settings.Deadzone = (int)deadzoneUpDown.Maximum;
            if (settings.ChannelLabelWidth < 80) settings.ChannelLabelWidth = 80;
            if (settings.ChannelLabelWidth > 240) settings.ChannelLabelWidth = 240;
            if (settings.ChannelRowHeight < 40) settings.ChannelRowHeight = 40;
            if (settings.ChannelRowHeight > 120) settings.ChannelRowHeight = 120;

            isApplyingSettings = true;
            channelCountUpDown.Value = settings.ChannelCount;
            deadzoneUpDown.Value = settings.Deadzone;
            channelRowHeightUpDown.Value = settings.ChannelRowHeight;
            channelLabelWidthUpDown.Value = settings.ChannelLabelWidth;
            manualPortCheckBox.Checked = settings.ManualPortEnabled;
            isApplyingSettings = false;

            PopulatePortList();
            UpdateManualPortUi();
            BuildChannelRows(settings.ChannelCount);
        }

        private void ApplyChannelSelections()
        {
            for (var i = 0; i < channelRows.Count; i++)
            {
                EnsureChannelSettings(i);
                var deviceId = settings.Channels[i].DeviceId;
                var match = audioDevices.FirstOrDefault(d => d.Id == deviceId);

                if (match != null)
                {
                    channelRows[i].DeviceComboBox.SelectedItem = match;
                }
                else if (audioDevices.Count > 0)
                {
                    channelRows[i].DeviceComboBox.SelectedIndex = 0;
                    settings.Channels[i].DeviceId = audioDevices[0].Id;
                }

                UpdateDeviceToolTip(channelRows[i].DeviceComboBox);
                UpdateComboDropDownWidth(channelRows[i].DeviceComboBox);
            }
        }

        private void EnsureChannelSettings(int index)
        {
            while (settings.Channels.Count <= index)
            {
                settings.Channels.Add(new ChannelSettings());
            }
        }

        private void StartAutoScan()
        {
            if (settings.ManualPortEnabled)
            {
                UpdateStatus(StatusState.Idle, "Status: Manuelle Portwahl aktiv");
                return;
            }

            scanCts?.Cancel();
            scanCts = new CancellationTokenSource();
            var token = scanCts.Token;

            UpdateStatus(StatusState.Searching, "Status: Suche Mixer.");
            rescanButton.Enabled = false;

            Task.Run(async () =>
            {
                var port = await FindMixerPortAsync(token);
                if (token.IsCancellationRequested) return;

                BeginInvoke(() =>
                {
                    rescanButton.Enabled = true;
                    if (port == null)
                    {
                        UpdateStatus(StatusState.Idle, "Status: Kein Mixer gefunden");
                    }
                    else
                    {
                        ConnectToPort(port);
                    }
                });
            }, token);
        }

        // ROBUSTE Auto-Suche (mehr Baudraten, \n und \r\n, DTR/RTS, längere Wartezeit)
        private async Task<string?> FindMixerPortAsync(CancellationToken token)
        {
            var baudRates = new[] { 9600, 115200, 57600, 38400 };
            var newLines = new[] { "\n", "\r\n" };

            foreach (var portName in SerialPort.GetPortNames().OrderBy(p => p))
            {
                foreach (var baud in baudRates)
                {
                    foreach (var nl in newLines)
                    {
                        if (token.IsCancellationRequested) return null;

                        try
                        {
                            using var probe = new SerialPort(portName, baud)
                            {
                                ReadTimeout = 300,
                                WriteTimeout = 300,
                                NewLine = nl,
                                DtrEnable = true,
                                RtsEnable = true,
                            };

                            probe.Open();

                            // Viele Boards resetten beim Öffnen -> länger warten
                            await Task.Delay(1500, token);

                            // Boot/Debug wegwerfen
                            probe.DiscardInBuffer();
                            probe.DiscardOutBuffer();

                            // Handshake
                            probe.WriteLine("HELLO_MIXER");

                            var deadline = DateTime.UtcNow.AddMilliseconds(1500);
                            while (DateTime.UtcNow < deadline && !token.IsCancellationRequested)
                            {
                                try
                                {
                                    var line = probe.ReadLine()?.Trim();
                                    if (string.IsNullOrEmpty(line)) continue;

                                    if (line.Equals("MIXER_READY", StringComparison.Ordinal))
                                        return portName;
                                }
                                catch (TimeoutException)
                                {
                                    // weiter probieren bis deadline
                                }
                            }
                        }
                        catch
                        {
                            // Port busy / falsche Settings / kein Device -> nächster Versuch
                        }
                    }
                }
            }

            return null;
        }

        private void ConnectToPort(string portName)
        {
            CloseSerialPort();

            try
            {
                serialPort = new SerialPort(portName, 9600)
                {
                    ReadTimeout = 500,
                    WriteTimeout = 500,
                    NewLine = "\n",        // \n funktioniert auch bei \r\n (Trim entfernt \r)
                    DtrEnable = true,
                    RtsEnable = true,
                };

                serialPort.DataReceived += SerialPortOnDataReceived;
                serialPort.Open();

                UpdateStatus(StatusState.Connected, $"Status: Verbunden ({portName})");
            }
            catch (Exception ex)
            {
                UpdateStatus(StatusState.Error, $"Status: Fehler beim Verbinden ({ex.Message})");
            }
        }

        private void CloseSerialPort()
        {
            if (serialPort == null) return;

            try
            {
                serialPort.DataReceived -= SerialPortOnDataReceived;
                serialPort.Close();
            }
            catch
            {
            }
            finally
            {
                serialPort.Dispose();
                serialPort = null;
            }
        }

        private void SerialPortOnDataReceived(object? sender, SerialDataReceivedEventArgs e)
        {
            if (serialPort == null) return;

            try
            {
                var line = serialPort.ReadLine();
                HandleMixerLine(line);
            }
            catch (TimeoutException)
            {
            }
            catch
            {
            }
        }

        private void HandleMixerLine(string line)
        {
            var parts = line.Split('|', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return;

            BeginInvoke(() =>
            {
                for (var i = 0; i < parts.Length && i < channelRows.Count; i++)
                {
                    if (!int.TryParse(parts[i], out var value))
                        continue;

                    if (value < 0) value = 0;
                    if (value > 1023) value = 1023;

                    if (lastValues.Length > i && lastValues[i] != -1)
                    {
                        if (Math.Abs(lastValues[i] - value) < settings.Deadzone)
                            continue;
                    }

                    lastValues[i] = value;
                    UpdateLevelBar(channelRows[i], value);
                    channelRows[i].LevelPercentLabel.Text = $"{Math.Round(value / 1023f * 100):0}%";

                    var deviceId = settings.Channels.ElementAtOrDefault(i)?.DeviceId;
                    if (!string.IsNullOrWhiteSpace(deviceId))
                    {
                        if (channelRows[i].IsMuted)
                        {
                            audioManager.SetDeviceVolume(deviceId, 0f);
                        }
                        else
                        {
                            var volume = value / 1023f;
                            audioManager.SetDeviceVolume(deviceId, volume);
                        }
                    }
                }
            });
        }

        private void SaveSettingsToFile()
        {
            using var dialog = new SaveFileDialog
            {
                Filter = "Mixer Einstellungen (*.json)|*.json",
                FileName = "mixer-settings.json",
            };

            if (dialog.ShowDialog(this) != DialogResult.OK) return;

            var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(dialog.FileName, json);
        }

        private void LoadSettingsFromFile()
        {
            using var dialog = new OpenFileDialog
            {
                Filter = "Mixer Einstellungen (*.json)|*.json",
            };

            if (dialog.ShowDialog(this) != DialogResult.OK) return;

            if (!TryReadSettingsFromPath(dialog.FileName, true, out var loaded) || loaded == null)
            {
                return;
            }

            settings = loaded;
            RebuildUi();
            PersistLastConfigIdentifier(dialog.FileName);
        }

        private static void EnableDoubleBuffering(Control control)
        {
            typeof(Control).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)
                ?.SetValue(control, true, null);
        }

        // Autoload/LastConfig: Pfad/Identifier wird in %AppData%\Audio_Mixer\last_config.txt gespeichert und beim Start geladen.
        private static string GetLastConfigPath()
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Audio_Mixer");
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, "last_config.txt");
        }

        private void PersistLastConfigIdentifier(string identifier)
        {
            try
            {
                File.WriteAllText(GetLastConfigPath(), identifier);
            }
            catch
            {
            }
        }

        private bool TryAutoLoadLastConfig()
        {
            string? identifier = null;
            try
            {
                var path = GetLastConfigPath();
                if (File.Exists(path))
                {
                    identifier = File.ReadAllText(path)?.Trim();
                }
            }
            catch
            {
            }

            if (string.IsNullOrWhiteSpace(identifier))
            {
                return false;
            }

            if (identifier.StartsWith("profile:", StringComparison.OrdinalIgnoreCase))
            {
                var profileName = identifier.Substring("profile:".Length);
                LoadProfile(profileName, false);
                return true;
            }

            if (!TryReadSettingsFromPath(identifier, false, out var loaded))
            {
                return false;
            }

            return ApplyLoadedSettings(loaded);
        }

        private bool TryReadSettingsFromPath(string path, bool showReadError, out MixerSettings? loaded)
        {
            loaded = null;
            if (!File.Exists(path))
            {
                if (showReadError)
                {
                    MessageBox.Show(this, "Die Datei konnte nicht gefunden werden.", "Fehler",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false;
            }

            string json;
            try
            {
                json = File.ReadAllText(path);
            }
            catch (Exception ex)
            {
                if (showReadError)
                {
                    MessageBox.Show(this, $"Einstellungen konnten nicht gelesen werden: {ex.Message}", "Fehler",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false;
            }

            loaded = MixerSettings.LoadBestEffort(json, MaxChannels);
            return true;
        }

        private bool ApplyLoadedSettings(MixerSettings? loaded)
        {
            if (loaded == null)
            {
                return false;
            }

            settings = loaded;
            RebuildUi();
            return true;
        }

        private sealed class ChannelRow
        {
            public ChannelRow(int index, ComboBox deviceComboBox, Panel levelTrack, Panel levelFill,
                Label levelPercentLabel, Button muteButton)
            {
                Index = index;
                DeviceComboBox = deviceComboBox;
                LevelTrack = levelTrack;
                LevelFill = levelFill;
                LevelPercentLabel = levelPercentLabel;
                MuteButton = muteButton;
            }

            public int Index { get; }
            public ComboBox DeviceComboBox { get; }
            public Panel LevelTrack { get; }
            public Panel LevelFill { get; }
            public Label LevelPercentLabel { get; }
            public Button MuteButton { get; }
            public int LevelValue { get; set; }
            public bool IsMuted { get; set; }
        }

        private enum StatusState
        {
            Idle,
            Searching,
            Connected,
            Error,
        }

        private enum ViewKind
        {
            Mixer,
            Settings,
        }

        private void UpdateLevelBar(ChannelRow row, int value)
        {
            row.LevelValue = value;
            var trackWidth = row.LevelTrack.ClientSize.Width - row.LevelTrack.Padding.Horizontal;
            trackWidth = Math.Max(0, trackWidth);
            var fillWidth = (int)Math.Round(trackWidth * value / 1023f);
            row.LevelFill.Width = Math.Max(0, fillWidth);
        }

        private sealed class StatusDot : Control
        {
            public Color DotColor { get; set; } = Color.DimGray;

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using var brush = new SolidBrush(DotColor);
                e.Graphics.FillEllipse(brush, 0, 0, Width - 1, Height - 1);
                using var pen = new Pen(Color.Black);
                e.Graphics.DrawEllipse(pen, 0, 0, Width - 1, Height - 1);
            }

            protected override void OnResize(EventArgs e)
            {
                base.OnResize(e);
                Invalidate();
            }
        }

        private sealed class AudioDeviceItem
        {
            public AudioDeviceItem(string id, string name)
            {
                Id = id;
                Name = name;
            }

            public string Id { get; }
            public string Name { get; }

            public override string ToString() => Name;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
