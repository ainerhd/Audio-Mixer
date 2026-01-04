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

        private TabControl mainTabControl = null!;
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

        private Color BackgroundColor => Color.FromArgb(settings.BackgroundColorArgb);
        private Color SurfaceColor => Color.FromArgb(settings.SurfaceColorArgb);
        private Color SurfaceAccentColor => Color.FromArgb(settings.SurfaceAccentColorArgb);
        private Color AccentColor => Color.FromArgb(settings.AccentColorArgb);
        private Color MutedTextColor => Color.FromArgb(settings.MutedTextColorArgb);

        private readonly ToolTip deviceToolTip = new() { ShowAlways = true };
        private StatusState statusState = StatusState.Idle;

        public Form1()
        {
            InitializeComponent();
            BuildUi();
            UpdateDevices();
            ApplySettings(settings);
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
            BackColor = BackgroundColor;
            ForeColor = Color.White;
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);

            mainTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                BackColor = BackgroundColor,
            };
            Controls.Add(mainTabControl);

            var mixerTab = new TabPage("Mixer")
            {
                BackColor = BackgroundColor,
                ForeColor = Color.White,
            };
            mainTabControl.TabPages.Add(mixerTab);

            var settingsTab = new TabPage("Einstellungen")
            {
                BackColor = BackgroundColor,
                ForeColor = Color.White,
            };
            mainTabControl.TabPages.Add(settingsTab);

            BuildMixerTabContent(mixerTab);
            BuildSettingsTabContent(settingsTab);
        }

        private void BuildMixerTabContent(TabPage tab)
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
                ForeColor = MutedTextColor,
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

            var gamingButton = CreateHeaderButton("Gaming");
            gamingButton.Click += (_, _) => LoadProfile("gaming");
            actionPanel.Controls.Add(gamingButton);

            var streamingButton = CreateHeaderButton("Streaming");
            streamingButton.Click += (_, _) => LoadProfile("streaming");
            actionPanel.Controls.Add(streamingButton);

            var officeButton = CreateHeaderButton("Office");
            officeButton.Click += (_, _) => LoadProfile("office");
            actionPanel.Controls.Add(officeButton);

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
            };
            channelsContainer.Controls.Add(channelsHeader, 0, 0);

            channelsHeaderTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 3,
                AutoSize = true,
            };
            ApplyChannelColumnStyles(channelsHeaderTable);
            channelsHeaderTable.Controls.Add(CreateHeaderLabel("Kanal"), 0, 0);
            channelsHeaderTable.Controls.Add(CreateHeaderLabel("Audio-Ausgang"), 1, 0);
            channelsHeaderTable.Controls.Add(CreateHeaderLabel("Level"), 2, 0);
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
                ColumnCount = 3,
                AutoSize = true,
            };
            ApplyChannelColumnStyles(channelsTable);
            channelsScrollPanel.Controls.Add(channelsTable);
        }

        private void BuildSettingsTabContent(TabPage tab)
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
                RowCount = 2,
                AutoSize = true,
            };
            generalContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            generalContainer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            generalCard.Controls.Add(generalContainer);

            generalContainer.Controls.Add(CreateSectionHeader("Allgemeine Einstellungen"), 0, 0);

            var settingsPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 6,
                AutoSize = true,
            };
            for (var i = 0; i < settingsPanel.ColumnCount; i++)
            {
                settingsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16.6f));
            }
            generalContainer.Controls.Add(settingsPanel, 0, 1);

            settingsPanel.Controls.Add(CreateFieldLabel("Kanäle:"), 0, 0);
            channelCountUpDown = new NumericUpDown
            {
                Minimum = 1,
                Maximum = MaxChannels,
                Value = settings.ChannelCount,
                Dock = DockStyle.Fill,
                BackColor = SurfaceAccentColor,
                ForeColor = Color.White,
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
                BackColor = SurfaceAccentColor,
                ForeColor = Color.White,
            };
            deadzoneUpDown.ValueChanged += (_, _) =>
            {
                if (isApplyingSettings) return;
                settings.Deadzone = (int)deadzoneUpDown.Value;
            };
            settingsPanel.Controls.Add(deadzoneUpDown, 3, 0);

            loadSettingsButton = new Button
            {
                Text = "Laden",
                Dock = DockStyle.Fill,
                BackColor = SurfaceAccentColor,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
            };
            loadSettingsButton.FlatAppearance.BorderSize = 0;
            loadSettingsButton.Click += (_, _) => LoadSettingsFromFile();
            settingsPanel.Controls.Add(loadSettingsButton, 4, 0);

            saveSettingsButton = new Button
            {
                Text = "Speichern",
                Dock = DockStyle.Fill,
                BackColor = SurfaceAccentColor,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
            };
            saveSettingsButton.FlatAppearance.BorderSize = 0;
            saveSettingsButton.Click += (_, _) => SaveSettingsToFile();
            settingsPanel.Controls.Add(saveSettingsButton, 5, 0);

            var connectionCard = CreateCardPanel();
            connectionCard.Margin = new Padding(0, 0, 0, 16);
            root.Controls.Add(connectionCard);

            var connectionContainer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                AutoSize = true,
            };
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
            connectionPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            connectionPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            connectionPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
            connectionPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
            connectionContainer.Controls.Add(connectionPanel, 0, 1);

            manualPortCheckBox = new CheckBox
            {
                Text = "Manuelle Portwahl aktivieren",
                Dock = DockStyle.Fill,
                ForeColor = Color.White,
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
                BackColor = SurfaceAccentColor,
                ForeColor = Color.White,
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
                BackColor = SurfaceAccentColor,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
            };
            refreshPortsButton.FlatAppearance.BorderSize = 0;
            refreshPortsButton.Click += (_, _) => PopulatePortList();
            connectionPanel.Controls.Add(refreshPortsButton, 2, 0);

            manualConnectButton = new Button
            {
                Text = "Verbinden",
                Dock = DockStyle.Fill,
                BackColor = AccentColor,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
            };
            manualConnectButton.FlatAppearance.BorderSize = 0;
            manualConnectButton.Click += (_, _) => ConnectManualPort();
            connectionPanel.Controls.Add(manualConnectButton, 3, 0);

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

            AddColorPickerRow(colorPanel, 0, "Hintergrund", () => BackgroundColor,
                color => settings.BackgroundColorArgb = color.ToArgb());
            AddColorPickerRow(colorPanel, 1, "Kartenfläche", () => SurfaceColor,
                color => settings.SurfaceColorArgb = color.ToArgb());
            AddColorPickerRow(colorPanel, 2, "Kartenakzent", () => SurfaceAccentColor,
                color => settings.SurfaceAccentColorArgb = color.ToArgb());
            AddColorPickerRow(colorPanel, 3, "Akzent", () => AccentColor,
                color => settings.AccentColorArgb = color.ToArgb());
            AddColorPickerRow(colorPanel, 4, "Gedämpfter Text", () => MutedTextColor,
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
                BackColor = SurfaceAccentColor,
                ForeColor = Color.White,
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
                BackColor = SurfaceAccentColor,
                ForeColor = Color.White,
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
                ForeColor = MutedTextColor,
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
                ForeColor = Color.White,
                AutoSize = true,
            };
        }

        private Panel CreateCardPanel()
        {
            return new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = SurfaceColor,
                Padding = new Padding(16),
            };
        }

        private Button CreateHeaderButton(string text)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = true,
                BackColor = AccentColor,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                Margin = new Padding(4, 0, 0, 0),
                Padding = new Padding(8, 4, 8, 4),
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private void ApplyChannelColumnStyles(TableLayoutPanel table)
        {
            table.ColumnStyles.Clear();
            table.ColumnCount = 3;
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, settings.ChannelLabelWidth));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
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
            row.MuteButton.BackColor = row.IsMuted ? AccentColor : SurfaceAccentColor;

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

        private void LoadProfile(string profileName)
        {
            var profilePath = Path.Combine(AppContext.BaseDirectory, $"{profileName}.json");
            MixerSettings? profile = null;

            if (File.Exists(profilePath))
            {
                try
                {
                    var json = File.ReadAllText(profilePath);
                    profile = JsonSerializer.Deserialize<MixerSettings>(json);
                }
                catch
                {
                    profile = null;
                }
            }

            profile ??= GetFallbackProfile(profileName);
            settings = profile;
            RebuildUi();
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
                ForeColor = MutedTextColor,
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
            };
            var button = new Button
            {
                Text = "Farbe wählen",
                Dock = DockStyle.Fill,
                BackColor = SurfaceAccentColor,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
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
            var selectedTabIndex = mainTabControl?.SelectedIndex ?? 0;
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

            if (mainTabControl != null && selectedTabIndex < mainTabControl.TabCount)
            {
                mainTabControl.SelectedIndex = selectedTabIndex;
            }

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
                var rowBackground = i % 2 == 0 ? SurfaceAccentColor : SurfaceColor;
                channelsTable.RowStyles.Add(new RowStyle(SizeType.Absolute, settings.ChannelRowHeight));

                var label = new Label
                {
                    Text = $"Kanal {i + 1}",
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleLeft,
                };
                channelsTable.Controls.Add(CreateCellPanel(label, rowBackground), 0, rowIndex);

                var deviceCombo = new ComboBox
                {
                    Dock = DockStyle.Fill,
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    BackColor = SurfaceColor,
                    ForeColor = Color.White,
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

                var progress = new ProgressBar
                {
                    Dock = DockStyle.Fill,
                    Maximum = 1023,
                    Value = 0,
                    Style = ProgressBarStyle.Continuous,
                };
                var muteButton = new Button
                {
                    Text = "Mute",
                    Dock = DockStyle.Fill,
                    BackColor = SurfaceAccentColor,
                    FlatStyle = FlatStyle.Flat,
                    ForeColor = Color.White,
                    Margin = new Padding(6, 2, 0, 2),
                };
                muteButton.FlatAppearance.BorderSize = 0;

                var levelContainer = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 2,
                    RowCount = 1,
                };
                levelContainer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                levelContainer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 64));
                levelContainer.Controls.Add(progress, 0, 0);
                levelContainer.Controls.Add(muteButton, 1, 0);
                channelsTable.Controls.Add(CreateCellPanel(levelContainer, rowBackground), 2, rowIndex);

                var row = new ChannelRow(i, deviceCombo, progress, muteButton);
                muteButton.Click += (_, _) => ToggleMute(row);
                channelRows.Add(row);
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
                    channelRows[i].LevelBar.Value = value;

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

            try
            {
                var json = File.ReadAllText(dialog.FileName);
                var loaded = JsonSerializer.Deserialize<MixerSettings>(json);
                if (loaded != null)
                {
                    settings = loaded;
                    RebuildUi();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Einstellungen konnten nicht geladen werden: {ex.Message}", "Fehler",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private sealed class ChannelRow
        {
            public ChannelRow(int index, ComboBox deviceComboBox, ProgressBar levelBar, Button muteButton)
            {
                Index = index;
                DeviceComboBox = deviceComboBox;
                LevelBar = levelBar;
                MuteButton = muteButton;
            }

            public int Index { get; }
            public ComboBox DeviceComboBox { get; }
            public ProgressBar LevelBar { get; }
            public Button MuteButton { get; }
            public bool IsMuted { get; set; }
        }

        private enum StatusState
        {
            Idle,
            Searching,
            Connected,
            Error,
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
