// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace NcTalkOutlookAddIn.Controllers
{
    internal sealed class MailBodyInsertionController
    {
        private readonly NextcloudTalkAddIn _owner;

        private readonly MailInteropController _mailInteropController;

        internal MailBodyInsertionController(NextcloudTalkAddIn owner, MailInteropController mailInteropController)
        {
            _owner = owner;
            _mailInteropController = mailInteropController;
        }

        internal bool InsertHtmlIntoMail(Outlook.MailItem mail, string html)
        {
            if (mail == null || string.IsNullOrWhiteSpace(html))
            {
                return false;
            }
            if (_mailInteropController.IsActiveInlineResponse(mail))
            {
                if (TryInsertHtmlIntoActiveInlineResponseWordEditor(mail, html))
                {
                    DiagnosticsLogger.Log(LogCategories.Core, "Inserted HTML block into inline response (ActiveInlineResponseWordEditor).");
                    return true;
                }

                DiagnosticsLogger.Log(LogCategories.Core, "Failed to insert HTML into inline response: inline WordEditor insertion failed.");
                MessageBox.Show(
                    string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, "inline WordEditor insertion failed"),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
            if (TryInsertHtmlIntoInspectorWordEditor(mail, html))
            {
                DiagnosticsLogger.Log(LogCategories.Core, "Inserted HTML block into mail (WordEditor InsertFile primary).");
                return true;
            }

            if (TryInsertHtmlIntoMailBody(mail, html))
            {
                DiagnosticsLogger.Log(LogCategories.Core, "Inserted HTML block into mail (HTMLBody compatibility fallback).");
                return true;
            }

            DiagnosticsLogger.Log(LogCategories.Core, "Failed to insert HTML into mail: all insertion paths exhausted.");
            MessageBox.Show(
                string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, "all insertion paths exhausted"),
                Strings.DialogTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return false;
        }

        internal bool InsertPlainTextIntoMail(Outlook.MailItem mail, string plainText)
        {
            if (mail == null || string.IsNullOrWhiteSpace(plainText))
            {
                return false;
            }

            try
            {
                string insertText = PlainTextUtilities.NormalizeCrLfAndTrim(plainText);
                if (!string.IsNullOrEmpty(insertText))
                {
                    insertText += "\r\n\r\n";
                }

                string source;
                if (TryInsertPlainTextViaWordEditor(mail, insertText, out source))
                {
                    DiagnosticsLogger.Log(LogCategories.Core, "Inserted plain-text share block into mail (source=" + source + ").");
                    return true;
                }

                throw new InvalidOperationException("Outlook WordEditor insertion point unavailable.");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert plain-text share block into mail.", ex);
                MessageBox.Show(
                    string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, ex.Message),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }

        private bool TryInsertPlainTextViaWordEditor(Outlook.MailItem mail, string text, out string source)
        {
            source = string.Empty;
            if (mail == null || string.IsNullOrEmpty(text))
            {
                return false;
            }

            Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;
            bool activeInlineMatches = _mailInteropController.IsActiveInlineResponse(mail);
            OutlookWordEditorContext context;

            try
            {
                if (activeInlineMatches)
                {
                    if (!OutlookWordEditorContext.TryOpenInline(application, mail, "plain-text share insert", null, out context))
                    {
                        return false;
                    }
                }
                else if (!OutlookWordEditorContext.TryOpenInspector(mail, "plain-text share insert", out context))
                {
                    return false;
                }

                using (context)
                {
                    source = context.Source ?? string.Empty;
                    if (activeInlineMatches)
                    {
                        int replyCursorStart = context.GetDocumentStart(0);
                        context.SetSelectionRange(replyCursorStart, replyCursorStart);
                        context.TypeParagraph();
                        context.TypeParagraph();
                        context.TypeText(text);
                        context.SetSelectionRange(replyCursorStart, replyCursorStart);
                    }
                    else
                    {
                        context.TypeText(text);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert plain text via Outlook WordEditor.", ex);
                return false;
            }
        }

        internal static bool IsPlainTextMail(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                return false;
            }

            try
            {
                return mail.BodyFormat == Outlook.OlBodyFormat.olFormatPlain;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read MailItem.BodyFormat.", ex);
                return false;
            }
        }

        private bool TryInsertHtmlIntoActiveInlineResponseWordEditor(Outlook.MailItem mail, string html)
        {
            Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;
            OutlookWordEditorContext context;
            string tempHtmlPath = null;

            try
            {
                if (!OutlookWordEditorContext.TryOpenInline(application, mail, "inline share insert", null, out context))
                {
                    return false;
                }

                using (context)
                {
                    InsertHtmlIntoWordEditor(context, html, "nc4ol-inline-share-", ref tempHtmlPath);
                }

                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert HTML into inline response Word editor.", ex);
                return false;
            }
            finally
            {
                WordHtmlInsertionFile.TryDeleteTemporaryHtmlFile(tempHtmlPath, "Failed to delete temporary inline share HTML file.");
            }
        }

        private static bool TryInsertHtmlIntoInspectorWordEditor(Outlook.MailItem mail, string html)
        {
            OutlookWordEditorContext context;
            string tempHtmlPath = null;

            try
            {
                if (!OutlookWordEditorContext.TryOpenInspector(mail, "HTML share insert", out context))
                {
                    return false;
                }

                using (context)
                {
                    InsertHtmlIntoWordEditor(context, html, "nc4ol-share-", ref tempHtmlPath);
                }

                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert HTML via inspector WordEditor.", ex);
                return false;
            }
            finally
            {
                WordHtmlInsertionFile.TryDeleteTemporaryHtmlFile(tempHtmlPath, "Failed to delete temporary share HTML file.");
            }
        }

        private static void InsertHtmlIntoWordEditor(
            OutlookWordEditorContext context,
            string html,
            string tempFilePrefix,
            ref string tempHtmlPath)
        {
            int replyCursorStart = context.GetDocumentStart(0);
            context.SetSelectionRange(replyCursorStart, replyCursorStart);
            context.TypeParagraph();
            context.TypeParagraph();

            tempHtmlPath = Path.Combine(
                Path.GetTempPath(),
                tempFilePrefix + Guid.NewGuid().ToString("N") + ".html");
            File.WriteAllText(tempHtmlPath, WordHtmlInsertionFile.EnsureHtmlDocumentForWordInsert(html), new UTF8Encoding(true));
            context.InsertFile(tempHtmlPath);
            context.SetSelectionRange(replyCursorStart, replyCursorStart);
        }

        private static bool TryInsertHtmlIntoMailBody(Outlook.MailItem mail, string html)
        {
            try
            {
                string existing = mail.HTMLBody ?? string.Empty;
                string insertHtml = "<br><br>" + html;
                int bodyTagIndex = existing.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                if (bodyTagIndex >= 0)
                {
                    int bodyTagEnd = existing.IndexOf(">", bodyTagIndex);
                    if (bodyTagEnd >= 0)
                    {
                        mail.HTMLBody = existing.Insert(bodyTagEnd + 1, insertHtml);
                    }
                    else
                    {
                        mail.HTMLBody = insertHtml + existing;
                    }
                }
                else
                {
                    mail.HTMLBody = insertHtml + existing;
                }
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log(LogCategories.Core, "Failed to insert HTML via HTMLBody path: " + ex.Message);
                return false;
            }
        }

    }
}
