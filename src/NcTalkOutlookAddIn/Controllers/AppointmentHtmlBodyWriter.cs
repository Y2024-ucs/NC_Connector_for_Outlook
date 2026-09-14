// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    internal static class AppointmentHtmlBodyWriter
    {
        internal static bool TryWriteAppointmentHtmlBody(Outlook.AppointmentItem appointment, string html)
        {
            if (appointment == null || string.IsNullOrWhiteSpace(html))
            {
                return false;
            }

            Outlook.Application application = null;
            Outlook.MailItem stagingMail = null;

            try
            {
                application = appointment.Application;
                if (application == null)
                {
                    return false;
                }

                stagingMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                if (stagingMail == null)
                {
                    return false;
                }
                string bridgeHtml = HtmlTemplateSanitizer.PrepareTalkAppointmentHtmlForOutlookRtfBridge(html);
                if (string.IsNullOrWhiteSpace(bridgeHtml))
                {
                    bridgeHtml = html ?? string.Empty;
                }

                stagingMail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                stagingMail.HTMLBody = bridgeHtml;

                var rtfBody = stagingMail.RTFBody as byte[];
                if (rtfBody == null || rtfBody.Length == 0)
                {
                    DiagnosticsLogger.Log(LogCategories.Talk, "Appointment HTML->RTF bridge produced empty RTF body.");
                    return false;
                }

                appointment.RTFBody = rtfBody;
                DiagnosticsLogger.Log(LogCategories.Talk, "Appointment HTML body written via HTML->RTF bridge.");
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to write appointment HTML body via HTML->RTF bridge.", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(stagingMail, LogCategories.Talk, "Failed to release staging MailItem COM object.");
                ComInteropScope.TryRelease(application, LogCategories.Talk, "Failed to release Outlook application COM object for appointment HTML->RTF bridge.");
            }
        }
    }
}
