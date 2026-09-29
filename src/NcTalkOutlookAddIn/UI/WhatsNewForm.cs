// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Drawing;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    /// <summary>
    /// Shown once after an update: what is new, what to do, what the user gains.
    /// </summary>
    internal sealed class WhatsNewForm : ScaledForm
    {
        private readonly FlowLayoutPanel _content = new FlowLayoutPanel();

        internal WhatsNewForm(ReleaseNote note)
        {
            Text = "NC Connector " + note.Version + " – Was ist neu?";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(ScaleLogical(620), ScaleLogical(600));
            Icon = BrandingAssets.GetAppIcon(32);

            var okButton = new Button
            {
                Text = "Verstanden",
                DialogResult = DialogResult.OK,
                Size = new Size(ScaleLogical(120), ScaleLogical(32)),
                Location = new Point(ClientSize.Width - ScaleLogical(140), ClientSize.Height - ScaleLogical(48)),
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom
            };
            Controls.Add(okButton);
            AcceptButton = okButton;
            CancelButton = okButton;

            _content.FlowDirection = FlowDirection.TopDown;
            _content.WrapContents = false;
            _content.AutoScroll = true;
            _content.Location = new Point(ScaleLogical(20), ScaleLogical(68));
            _content.Size = new Size(ClientSize.Width - ScaleLogical(40), ClientSize.Height - ScaleLogical(68) - ScaleLogical(64));
            _content.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            Controls.Add(_content);

            var header = new BrandedHeader();
            BrandedHeader.AttachToParent(header, Controls, ScaleLogical(56));

            AddText("Ihr NC Connector wurde auf Version " + note.Version + " aktualisiert.", 3f, FontStyle.Bold, 0);
            AddSection("Das ist neu", note.WhatsNew);
            AddSection("Das müssen Sie tun", note.ToDo);
            AddSection("Ihre Vorteile", note.Benefits);

            UiThemeManager.ApplyToForm(this, null);
        }

        private void AddSection(string caption, string[] lines)
        {
            if (lines == null || lines.Length == 0) return;
            AddText(caption, 1.5f, FontStyle.Bold, ScaleLogical(14));
            foreach (string line in lines)
            {
                AddText("•  " + line, 0f, FontStyle.Regular, ScaleLogical(4));
            }
        }

        private void AddText(string value, float extraSize, FontStyle style, int topMargin)
        {
            int width = _content.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - ScaleLogical(8);
            _content.Controls.Add(new Label
            {
                Text = value,
                AutoSize = true,
                MaximumSize = new Size(width, 0),
                Font = new Font(Font.FontFamily, Font.Size + extraSize, style),
                Margin = new Padding(0, topMargin, 0, 0),
                UseMnemonic = false
            });
        }
    }
}
