// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.UI
{
    internal sealed partial class SettingsForm
    {
        private void ApplySignatureTabLayout()
        {
            int groupLeft = ScaleLogical(24);
            int groupTop = ScaleLogical(20);
            int groupWidth = Math.Max(ScaleLogical(360), _signatureTab.ClientSize.Width - (groupLeft * 2));
            _signatureDefaultsGroup.SetBounds(groupLeft, groupTop, groupWidth, _signatureDefaultsGroup.Height);

            int innerPadding = ScaleLogical(12);
            int contentWidth = Math.Max(ScaleLogical(180), _signatureDefaultsGroup.ClientSize.Width - (innerPadding * 2));
            int y = ScaleLogical(26);
            int checkGap = ScaleLogical(8);

            _signatureIdentityHintLabel.Location = new Point(innerPadding, y);
            _signatureIdentityHintLabel.MaximumSize = new Size(contentWidth, 0);
            _signatureIdentityHintLabel.AutoSize = true;
            y = _signatureIdentityHintLabel.Bottom + ScaleLogical(16);

            _emailSignatureOnComposeCheckBox.Location = new Point(innerPadding, y);
            y = _emailSignatureOnComposeCheckBox.Bottom + checkGap;

            _emailSignatureOnReplyCheckBox.Location = new Point(innerPadding + ScaleLogical(18), y);
            y = _emailSignatureOnReplyCheckBox.Bottom + checkGap;

            _emailSignatureOnForwardCheckBox.Location = new Point(innerPadding + ScaleLogical(18), y);
            y = _emailSignatureOnForwardCheckBox.Bottom + innerPadding;

            _signatureDefaultsGroup.Height = Math.Max(ScaleLogical(150), y);
        }

        private void InitializeSignatureTab()
        {
            _signatureTab.AutoScroll = true;
            _signatureTab.Padding = new Padding(12);

            _signatureDefaultsGroup.Text = Strings.SignatureDefaultsHeading;
            _signatureDefaultsGroup.Location = new Point(24, 20);
            _signatureDefaultsGroup.Size = new Size(480, 170);
            _signatureTab.Controls.Add(_signatureDefaultsGroup);

            _signatureIdentityHintLabel.Text = Strings.SignatureIdentityHint;
            _signatureIdentityHintLabel.ForeColor = Color.DimGray;
            _signatureIdentityHintLabel.AutoSize = true;
            _signatureDefaultsGroup.Controls.Add(_signatureIdentityHintLabel);

            _emailSignatureOnComposeCheckBox.Text = Strings.SignatureOnComposeLabel;
            _emailSignatureOnComposeCheckBox.AutoSize = true;
            _emailSignatureOnComposeCheckBox.CheckedChanged += (s, e) => UpdateControlState();
            _signatureDefaultsGroup.Controls.Add(_emailSignatureOnComposeCheckBox);

            _emailSignatureOnReplyCheckBox.Text = Strings.SignatureOnReplyLabel;
            _emailSignatureOnReplyCheckBox.AutoSize = true;
            _signatureDefaultsGroup.Controls.Add(_emailSignatureOnReplyCheckBox);

            _emailSignatureOnForwardCheckBox.Text = Strings.SignatureOnForwardLabel;
            _emailSignatureOnForwardCheckBox.AutoSize = true;
            _signatureDefaultsGroup.Controls.Add(_emailSignatureOnForwardCheckBox);

            ApplySignatureTabLayout();
        }

        private bool IsEmailSignaturePolicyAvailable()
        {
            return EmailSignaturePolicyService.IsAvailableForConfiguration(_backendPolicyStatus);
        }

        private string GetEmailSignatureUnavailableTooltip()
        {
            string entitlementTooltip = PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(_backendPolicyStatus);
            if (!string.IsNullOrWhiteSpace(entitlementTooltip))
            {
                return entitlementTooltip;
            }
            if (!PolicyUiHelper.IsPolicyDomainAvailable(_backendPolicyStatus, "email_signature"))
            {
                return Strings.SignatureBackendUpdateRequiredTooltip;
            }
            return Strings.SignatureBackendInactiveTooltip;
        }

    }
}
