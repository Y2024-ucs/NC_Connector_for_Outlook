// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Drawing;
using System.Windows.Forms;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class WarningPanelUiHelper
    {
        private static readonly Color WarningColor =
            Color.FromArgb(176, 0, 32);

        internal static void Initialize(
            Panel panel,
            Label titleLabel,
            Label textLabel,
            LinkLabel linkLabel,
            string title,
            string text,
            string linkText)
        {
            panel.Visible = false;
            panel.BackColor = Color.FromArgb(20, 176, 0, 32);
            panel.Paint += (sender, args) =>
            {
                ControlPaint.DrawBorder(
                    args.Graphics,
                    panel.ClientRectangle,
                    titleLabel.ForeColor,
                    ButtonBorderStyle.Solid);
            };

            titleLabel.AutoSize = true;
            titleLabel.ForeColor = WarningColor;
            titleLabel.Font = new Font(titleLabel.Font, FontStyle.Bold);
            titleLabel.Text = title ?? string.Empty;
            panel.Controls.Add(titleLabel);

            textLabel.AutoSize = true;
            textLabel.Text = text ?? string.Empty;
            panel.Controls.Add(textLabel);

            linkLabel.AutoSize = true;
            linkLabel.Text = linkText ?? string.Empty;
            linkLabel.LinkColor = Color.FromArgb(0, 130, 201);
            linkLabel.ActiveLinkColor = Color.FromArgb(0, 102, 153);
            linkLabel.VisitedLinkColor = Color.FromArgb(0, 130, 201);
            panel.Controls.Add(linkLabel);
        }

        internal static int Layout(
            Panel panel,
            Label titleLabel,
            Label textLabel,
            LinkLabel linkLabel,
            int left,
            int top,
            int width,
            int padding,
            int minimumTextWidth,
            int titleTextGap,
            int textLinkGap)
        {
            if (!panel.Visible)
            {
                panel.SetBounds(left, top, width, 0);
                return 0;
            }

            int textWidth = System.Math.Max(
                minimumTextWidth,
                width - (padding * 2));
            titleLabel.Location = new Point(padding, padding);
            titleLabel.MaximumSize = new Size(textWidth, 0);

            int textTop = titleLabel.Bottom + titleTextGap;
            textLabel.Location = new Point(padding, textTop);
            textLabel.MaximumSize = new Size(textWidth, 0);

            int linkTop = textLabel.Bottom + textLinkGap;
            linkLabel.Location = new Point(padding, linkTop);
            linkLabel.MaximumSize = new Size(textWidth, 0);

            int height = (linkLabel.Visible ? linkLabel.Bottom : textLabel.Bottom) + padding;
            panel.SetBounds(left, top, width, height);
            return height;
        }
    }
}
