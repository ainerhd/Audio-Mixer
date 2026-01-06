using System.Drawing;
using System.Windows.Forms;

namespace Audio_Mixer.UI
{
    public static class SettingsLayoutHelper
    {
        public static Panel EnsureScrollableTabPage(TabPage page)
        {
            var scrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
            };
            page.Controls.Add(scrollPanel);
            return scrollPanel;
        }

        public static TableLayoutPanel BuildStandardGrid(TabPage page)
        {
            var scrollPanel = EnsureScrollableTabPage(page);
            return BuildStandardGrid(scrollPanel, 2, new Padding(0));
        }

        public static TableLayoutPanel BuildStandardGrid(Control parent, int columnCount, Padding? margin = null)
        {
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = columnCount,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = margin ?? new Padding(0, 8, 0, 0),
                GrowStyle = TableLayoutPanelGrowStyle.AddRows,
            };

            grid.ColumnStyles.Clear();
            if (columnCount == 2)
            {
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            }
            else
            {
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            }

            parent.Controls.Add(grid);
            return grid;
        }

        public static void AddRow(TableLayoutPanel grid, string labelText, Control control, Control? optionalControl = null)
        {
            var label = new Label
            {
                Text = labelText,
                AutoSize = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
            };
            AddRow(grid, label, control, optionalControl);
        }

        public static void AddRow(TableLayoutPanel grid, Label label, Control control, Control? optionalControl = null)
        {
            var rowIndex = grid.RowCount;
            grid.RowCount += 1;
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            label.Margin = new Padding(0, 6, 12, 6);
            control.Margin = new Padding(0, 6, 0, 6);
            control.Dock = DockStyle.Fill;

            grid.Controls.Add(label, 0, rowIndex);
            grid.Controls.Add(control, 1, rowIndex);

            if (optionalControl != null && grid.ColumnCount >= 3)
            {
                optionalControl.Margin = new Padding(8, 6, 0, 6);
                optionalControl.Dock = DockStyle.Fill;
                grid.Controls.Add(optionalControl, 2, rowIndex);
                return;
            }

            if (grid.ColumnCount >= 3)
            {
                TableLayoutPanel.SetColumnSpan(control, 2);
            }
        }
    }
}
