// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    internal sealed class SeparatePasswordDeliveryController
    {
        internal const string SecretDeliveryPlaceholder =
            "__NCC_SECRET_DELIVERY_VALUE__";

        private readonly NextcloudTalkAddIn _owner;
        private readonly ManagedEmailSignatureController _passwordManagedEmailSignatureController;

        private sealed class SeparatePasswordSignatureSnapshot
        {
            internal bool Active { get; set; }

            internal string Reason { get; set; }

            internal string UserEmail { get; set; }

            internal string Html { get; set; }

            internal string PlainText { get; set; }
        }

        internal SeparatePasswordDeliveryController(NextcloudTalkAddIn owner)
        {
            _owner = owner;
            _passwordManagedEmailSignatureController =
                new ManagedEmailSignatureController(owner);
        }

        internal void CaptureSeparatePasswordSignatureSnapshot(
            List<SeparatePasswordDispatchEntry> queue,
            BackendPolicyStatus policyStatus,
            AddinSettings settings,
            string composeKey)
        {
            if (queue == null || queue.Count == 0)
            {
                return;
            }
            SeparatePasswordSignatureSnapshot snapshot =
                BuildSeparatePasswordSignatureSnapshot(
                    policyStatus,
                    settings,
                    composeKey);
            foreach (SeparatePasswordDispatchEntry entry in queue)
            {
                if (entry == null)
                {
                    continue;
                }
                entry.SignatureActive = snapshot.Active;
                entry.SignatureUserEmail = snapshot.UserEmail;
                entry.SignatureHtml = snapshot.Html;
                entry.SignaturePlainText = snapshot.PlainText;
            }
        }

        internal void DispatchSeparatePasswordMailQueue(
            string composeKey,
            List<SeparatePasswordDispatchEntry> queue)
        {
            if (queue == null
                || queue.Count == 0
                || _owner.OutlookApplication == null)
            {
                return;
            }

            List<SeparatePasswordDispatchEntry> dispatches =
                ExpandSeparatePasswordDispatchEntries(queue);
            var sentRecipients = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            int attempted = 0;
            int sent = 0;
            int manual = 0;
            int secretsFallback = 0;

            foreach (SeparatePasswordDispatchEntry queued in dispatches)
            {
                SeparatePasswordDispatchEntry dispatch = null;
                Outlook.MailItem passwordMail = null;
                List<string> resolvedRecipients = null;
                try
                {
                    dispatch =
                        PrepareSeparatePasswordDispatch(
                            queued,
                            composeKey,
                            ref secretsFallback);
                    if (!IsDispatchUsable(dispatch))
                    {
                        continue;
                    }

                    attempted++;
                    passwordMail =
                        _owner.OutlookApplication.CreateItem(
                            Outlook.OlItemType.olMailItem)
                        as Outlook.MailItem;
                    if (passwordMail == null)
                    {
                        throw new InvalidOperationException(
                            "Password mail could not be created.");
                    }
                    resolvedRecipients = PopulatePasswordMail(
                        passwordMail,
                        dispatch,
                        composeKey);

                    NextcloudTalkAddIn.LogFileLinkMessage(
                        "Separate password mail direct send start (composeKey="
                        + (composeKey ?? string.Empty)
                        + ", recipients="
                        + resolvedRecipients.Count.ToString(
                            CultureInfo.InvariantCulture)
                        + ").");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Separate password mail build failed (composeKey="
                        + (composeKey ?? string.Empty)
                        + ").",
                        ex);
                    if (TryOpenSeparatePasswordFallback(
                            dispatch,
                            composeKey))
                    {
                        manual++;
                    }
                    else
                    {
                        ShowPasswordMailFailure(ex.Message);
                    }
                    ComInteropScope.TryRelease(
                        passwordMail,
                        LogCategories.FileLink,
                        "Failed to release incomplete password MailItem COM object.");
                    continue;
                }

                try
                {
                    ((Outlook._MailItem)passwordMail).Send();
                    sent++;
                    foreach (string recipient
                        in resolvedRecipients)
                    {
                        sentRecipients.Add(recipient);
                    }
                    NextcloudTalkAddIn.LogFileLinkMessage(
                        "Separate password mail direct send accepted (composeKey="
                        + (composeKey ?? string.Empty)
                        + ").");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Separate password mail direct send failed (composeKey="
                        + (composeKey ?? string.Empty)
                        + ").",
                        ex);
                    bool submittedOrAmbiguous =
                        ReadSubmittedOrAmbiguous(passwordMail);
                    if (!submittedOrAmbiguous
                        && TryOpenSeparatePasswordFallback(
                            dispatch,
                            composeKey))
                    {
                        manual++;
                    }
                    else if (!submittedOrAmbiguous)
                    {
                        ShowPasswordMailFailure(ex.Message);
                    }
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        passwordMail,
                        LogCategories.FileLink,
                        "Failed to release password MailItem COM object.");
                }
            }

            if (secretsFallback > 0
                && (sent > 0 || manual > 0))
            {
                ShowSecretsFallbackWarning();
            }
            if (attempted > 0 && sent == attempted)
            {
                _owner.ShowPasswordMailSuccessNotification(
                    sentRecipients.Count);
            }
            NextcloudTalkAddIn.LogFileLinkMessage(
                "Separate password direct dispatch completed (composeKey="
                + (composeKey ?? string.Empty)
                + ", attempted="
                + attempted.ToString(CultureInfo.InvariantCulture)
                + ", sent="
                + sent.ToString(CultureInfo.InvariantCulture)
                + ", manual="
                + manual.ToString(CultureInfo.InvariantCulture)
                + ").");
        }

        private List<string> PopulatePasswordMail(
            Outlook.MailItem mail,
            SeparatePasswordDispatchEntry dispatch,
            string composeKey)
        {
            if (mail == null || !IsDispatchUsable(dispatch))
            {
                throw new InvalidOperationException(
                    "Password mail is incomplete.");
            }
            mail.Subject =
                BuildSeparatePasswordMailSubject(dispatch);
            string effectiveSender =
                ApplyAndVerifySeparatePasswordSender(
                    mail,
                    dispatch,
                    composeKey,
                    true);
            ApplySeparatePasswordBody(mail, dispatch);
            ApplySeparatePasswordBackendSignature(
                mail,
                dispatch,
                effectiveSender,
                CapturedSignature(dispatch),
                composeKey);
            return ApplySeparatePasswordRecipientsForSend(
                mail,
                dispatch,
                composeKey);
        }

        private bool TryOpenSeparatePasswordFallback(
            SeparatePasswordDispatchEntry dispatch,
            string composeKey)
        {
            if (!IsDispatchUsable(dispatch)
                || _owner.OutlookApplication == null)
            {
                return false;
            }

            Outlook.MailItem fallback = null;
            try
            {
                fallback =
                    _owner.OutlookApplication.CreateItem(
                        Outlook.OlItemType.olMailItem)
                    as Outlook.MailItem;
                if (fallback == null)
                {
                    return false;
                }
                string toRecipients =
                    RecipientAddressList.BuildNormalizedRecipientCsv(dispatch.To);
                string ccRecipients =
                    RecipientAddressList.BuildNormalizedRecipientCsv(dispatch.Cc);
                string bccRecipients =
                    RecipientAddressList.BuildNormalizedRecipientCsv(dispatch.Bcc);
                if (RecipientAddressList.CountRecipientsInCsv(toRecipients)
                    + RecipientAddressList.CountRecipientsInCsv(ccRecipients)
                    + RecipientAddressList.CountRecipientsInCsv(bccRecipients) <= 0)
                {
                    throw new InvalidOperationException(
                        "Separate password fallback mail has no valid recipients.");
                }

                fallback.To = toRecipients;
                fallback.CC = ccRecipients;
                fallback.BCC = bccRecipients;
                fallback.Subject =
                    BuildSeparatePasswordMailSubject(dispatch);
                ApplyAndVerifySeparatePasswordSender(
                    fallback,
                    dispatch,
                    composeKey,
                    false);
                ApplySeparatePasswordBody(
                    fallback,
                    dispatch);
                fallback.Display(false);
                ApplySeparatePasswordBackendSignatureToDisplayedFallback(
                    fallback,
                    dispatch,
                    CapturedSignature(dispatch),
                    composeKey);
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password mail manual fallback opened (composeKey="
                    + (composeKey ?? string.Empty)
                    + ").");
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Separate password mail manual fallback failed (composeKey="
                    + (composeKey ?? string.Empty)
                    + ").",
                    ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(
                    fallback,
                    LogCategories.FileLink,
                    "Failed to release password fallback MailItem COM object.");
            }
        }

        private static bool ReadSubmittedOrAmbiguous(
            Outlook.MailItem mail)
        {
            try
            {
                return mail == null || mail.Submitted;
            }
            catch
            {
                return true;
            }
        }

        private static void ShowPasswordMailFailure(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }
            try
            {
                MessageBox.Show(
                    message.Trim(),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch
            {
                // The primary Outlook send must not fail because a warning
                // cannot be displayed.
            }
        }

        private static void ShowSecretsFallbackWarning()
        {
            try
            {
                MessageBox.Show(
                    Strings.SharingPasswordSecretsFallbackWarning,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch
            {
                // The primary Outlook send must not fail because a warning
                // cannot be displayed.
            }
        }

        private static SeparatePasswordSignatureSnapshot
            CapturedSignature(SeparatePasswordDispatchEntry entry)
        {
            return new SeparatePasswordSignatureSnapshot
            {
                Active = entry != null && entry.SignatureActive,
                Reason = "captured",
                UserEmail = entry != null
                    ? entry.SignatureUserEmail
                    : string.Empty,
                Html = entry != null ? entry.SignatureHtml : string.Empty,
                PlainText = entry != null
                    ? entry.SignaturePlainText
                    : string.Empty
            };
        }

        private static List<SeparatePasswordDispatchEntry> ExpandSeparatePasswordDispatchEntries(List<SeparatePasswordDispatchEntry> queue)
        {
            var expanded = new List<SeparatePasswordDispatchEntry>();
            if (queue == null)
            {
                return expanded;
            }

            for (int i = 0; i < queue.Count; i++)
            {
                SeparatePasswordDispatchEntry entry = queue[i];
                if (entry == null || entry.DeliveryMode != SharePasswordDeliveryMode.Secrets)
                {
                    expanded.Add(entry);
                    continue;
                }

                int added = 0;
                var seen =
                    new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                added += AddPerRecipientEntries(expanded, entry, RecipientAddressList.ExtractRecipientAddresses(entry.To), "to", seen);
                added += AddPerRecipientEntries(expanded, entry, RecipientAddressList.ExtractRecipientAddresses(entry.Cc), "cc", seen);
                added += AddPerRecipientEntries(expanded, entry, RecipientAddressList.ExtractRecipientAddresses(entry.Bcc), "bcc", seen);
                if (added == 0)
                {
                    expanded.Add(entry);
                }
            }

            return expanded;
        }

        private static int AddPerRecipientEntries(
            List<SeparatePasswordDispatchEntry> target,
            SeparatePasswordDispatchEntry source,
            List<string> recipients,
            string field,
            HashSet<string> seen)
        {
            if (target == null || source == null || recipients == null || recipients.Count == 0)
            {
                return 0;
            }

            int added = 0;
            for (int i = 0; i < recipients.Count; i++)
            {
                string address = RecipientAddressList.NormalizeRecipientAddress(recipients[i]);
                if (string.IsNullOrWhiteSpace(address))
                {
                    continue;
                }
                if (!seen.Add(address))
                {
                    continue;
                }
                SeparatePasswordDispatchEntry clone = ClonePasswordDispatch(source);
                clone.To = string.Equals(field, "to", StringComparison.OrdinalIgnoreCase) ? address : string.Empty;
                clone.Cc = string.Equals(field, "cc", StringComparison.OrdinalIgnoreCase) ? address : string.Empty;
                clone.Bcc = string.Equals(field, "bcc", StringComparison.OrdinalIgnoreCase) ? address : string.Empty;
                target.Add(clone);
                added++;
            }
            return added;
        }

        private SeparatePasswordDispatchEntry PrepareSeparatePasswordDispatch(
            SeparatePasswordDispatchEntry dispatch,
            string composeKey,
            ref int secretsFallbackCount)
        {
            if (dispatch == null || dispatch.DeliveryMode != SharePasswordDeliveryMode.Secrets)
            {
                return dispatch;
            }

            try
            {
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password Secrets link create start (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", to="
                    + RecipientAddressList.CountRecipientsInCsv(dispatch.To).ToString(CultureInfo.InvariantCulture)
                    + ", cc="
                    + RecipientAddressList.CountRecipientsInCsv(dispatch.Cc).ToString(CultureInfo.InvariantCulture)
                    + ", bcc="
                    + RecipientAddressList.CountRecipientsInCsv(dispatch.Bcc).ToString(CultureInfo.InvariantCulture)
                    + ", expireDays="
                    + dispatch.SecretsExpireDays.ToString(CultureInfo.InvariantCulture)
                    + ", mailPlainText="
                    + dispatch.IsPlainText.ToString(CultureInfo.InvariantCulture)
                    + ").");
                SecretsLinkResult secret = CreatePasswordSecret(dispatch);
                SeparatePasswordDispatchEntry prepared = BuildPasswordDispatchBody(
                    dispatch,
                    secret.ShareUrl,
                    SharePasswordDeliveryMode.Secrets,
                    secretLink: true);
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password Secrets link created (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", hasUuid="
                    + (!string.IsNullOrWhiteSpace(secret.Uuid)).ToString(CultureInfo.InvariantCulture)
                    + ", hasExpires="
                    + secret.Expires.HasValue.ToString(CultureInfo.InvariantCulture)
                    + ", hasHtml="
                    + (!string.IsNullOrWhiteSpace(prepared.Html)).ToString(CultureInfo.InvariantCulture)
                    + ", hasPlainText="
                    + (!string.IsNullOrWhiteSpace(prepared.PlainText)).ToString(CultureInfo.InvariantCulture)
                    + ").");
                return prepared;
            }
            catch (Exception ex)
            {
                secretsFallbackCount++;
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Separate password Secrets link creation failed, falling back to plain mail (composeKey=" + (composeKey ?? string.Empty) + ").",
                    ex);
                SeparatePasswordDispatchEntry fallback = ClonePasswordDispatch(dispatch);
                fallback.DeliveryMode = SharePasswordDeliveryMode.Plain;
                return fallback;
            }
        }

        private SecretsLinkResult CreatePasswordSecret(SeparatePasswordDispatchEntry dispatch)
        {
            TalkServiceConfiguration configuration =
                dispatch != null
                && dispatch.Origin != null
                && dispatch.Origin.IsComplete()
                    ? dispatch.Origin.ToConfiguration()
                    : null;
            if (configuration == null)
            {
                _owner.EnsureSettingsLoaded();
                if (_owner.CurrentSettings == null
                    || !_owner.SettingsAreComplete())
                {
                    throw new InvalidOperationException(
                        "Settings are incomplete.");
                }
                AddinSettings settings = _owner.CurrentSettings;
                configuration = new TalkServiceConfiguration(
                    settings.ServerUrl,
                    settings.Username,
                    settings.AppPassword);
            }

            var service = new SecretsService(configuration);
            return service.CreateSecretLink(
                dispatch.Password,
                BuildSecretsTitle(dispatch),
                dispatch.SecretsExpireDays);
        }

        private static string BuildSecretsTitle(SeparatePasswordDispatchEntry dispatch)
        {
            string shareLabel = dispatch != null ? (dispatch.ShareLabel ?? string.Empty).Trim() : string.Empty;
            if (string.IsNullOrWhiteSpace(shareLabel))
            {
                return "NCC share password";
            }
            return "NCC " + shareLabel;
        }

        private static SeparatePasswordDispatchEntry BuildPasswordDispatchBody(
            SeparatePasswordDispatchEntry source,
            string deliveryValue,
            SharePasswordDeliveryMode mode,
            bool secretLink)
        {
            FileLinkResult result = BuildFileLinkResultForDelivery(source, deliveryValue);
            string html = source.IsPlainText
                ? string.Empty
                : ReplaceDeliveryPlaceholder(
                    source.SecretsHtmlTemplate,
                    deliveryValue);
            if (!source.IsPlainText && string.IsNullOrWhiteSpace(html))
            {
                html = FileLinkHtmlBuilder.BuildPasswordOnly(result, source.LanguageOverride, source.BackendPolicyStatus, secretLink);
            }
            string plainText = source.IsPlainText
                ? ReplaceDeliveryPlaceholder(
                    source.SecretsPlainTextTemplate,
                    deliveryValue)
                : string.Empty;
            if (source.IsPlainText && string.IsNullOrWhiteSpace(plainText))
            {
                plainText = FileLinkHtmlBuilder.BuildPasswordOnlyPlainText(result, source.LanguageOverride, source.BackendPolicyStatus, secretLink);
            }

            SeparatePasswordDispatchEntry prepared = ClonePasswordDispatch(source);
            prepared.Password = deliveryValue ?? string.Empty;
            prepared.Html = html;
            prepared.PlainText = plainText;
            prepared.DeliveryMode = mode;
            return prepared;
        }

        private static string ReplaceDeliveryPlaceholder(
            string template,
            string value)
        {
            return string.IsNullOrWhiteSpace(template)
                ? string.Empty
                : template.Replace(
                    SecretDeliveryPlaceholder,
                    value ?? string.Empty);
        }

        private static FileLinkResult BuildFileLinkResultForDelivery(SeparatePasswordDispatchEntry dispatch, string deliveryValue)
        {
            if (dispatch == null)
            {
                throw new ArgumentNullException("dispatch");
            }

            return new FileLinkResult(
                dispatch.ShareUrl,
                dispatch.ShareId,
                dispatch.ShareToken,
                deliveryValue ?? string.Empty,
                dispatch.ExpireDate,
                dispatch.Permissions,
                dispatch.ShareLabel,
                dispatch.RelativePath);
        }

        private static SeparatePasswordDispatchEntry ClonePasswordDispatch(SeparatePasswordDispatchEntry source)
        {
            if (source == null)
            {
                return null;
            }

            return new SeparatePasswordDispatchEntry
            {
                ShareLabel = source.ShareLabel,
                ShareUrl = source.ShareUrl,
                ShareId = source.ShareId,
                ShareToken = source.ShareToken,
                RelativePath = source.RelativePath,
                ExpireDate = source.ExpireDate,
                Permissions = source.Permissions,
                Password = source.Password,
                Html = source.Html,
                PlainText = source.PlainText,
                SecretsHtmlTemplate = source.SecretsHtmlTemplate,
                SecretsPlainTextTemplate =
                    source.SecretsPlainTextTemplate,
                IsPlainText = source.IsPlainText,
                DeliveryMode = source.DeliveryMode,
                SecretsExpireDays = source.SecretsExpireDays,
                LanguageOverride = source.LanguageOverride,
                BackendPolicyStatus = source.BackendPolicyStatus,
                To = source.To,
                Cc = source.Cc,
                Bcc = source.Bcc,
                SenderEmail = source.SenderEmail,
                SendUsingAccountSmtpAddress = source.SendUsingAccountSmtpAddress,
                SentOnBehalfOfName = source.SentOnBehalfOfName,
                SignatureActive = source.SignatureActive,
                SignatureUserEmail = source.SignatureUserEmail,
                SignatureHtml = source.SignatureHtml,
                SignaturePlainText = source.SignaturePlainText,
                Origin = source.Origin != null
                    ? source.Origin.Clone()
                    : null
            };
        }

        private List<string> ApplySeparatePasswordRecipientsForSend(
            Outlook.MailItem mail,
            SeparatePasswordDispatchEntry dispatch,
            string composeKey)
        {
            if (mail == null)
            {
                throw new InvalidOperationException("Password mail is not available.");
            }

            List<string> toRecipients = RecipientAddressList.ExtractRecipientAddresses(dispatch != null ? dispatch.To : string.Empty);
            List<string> ccRecipients = RecipientAddressList.ExtractRecipientAddresses(dispatch != null ? dispatch.Cc : string.Empty);
            List<string> bccRecipients = RecipientAddressList.ExtractRecipientAddresses(dispatch != null ? dispatch.Bcc : string.Empty);
            int totalRecipients = toRecipients.Count + ccRecipients.Count + bccRecipients.Count;
            if (totalRecipients <= 0)
            {
                throw new InvalidOperationException("Separate password mail has no valid recipients.");
            }
            var resolvedRecipients = new List<string>();
            Outlook.Recipients recipients = null;
            try
            {
                recipients = mail.Recipients;
                if (recipients == null)
                {
                    throw new InvalidOperationException("Password mail recipients collection is not available.");
                }

                AddResolvedRecipients(recipients, toRecipients, Outlook.OlMailRecipientType.olTo, composeKey, resolvedRecipients);
                AddResolvedRecipients(recipients, ccRecipients, Outlook.OlMailRecipientType.olCC, composeKey, resolvedRecipients);
                AddResolvedRecipients(recipients, bccRecipients, Outlook.OlMailRecipientType.olBCC, composeKey, resolvedRecipients);

                bool resolvedAll = recipients.ResolveAll();
                if (!resolvedAll)
                {
                    throw new InvalidOperationException("Separate password mail recipients could not be resolved.");
                }
                if (resolvedRecipients.Count <= 0)
                {
                    throw new InvalidOperationException("Separate password mail has no resolvable recipients.");
                }
                return resolvedRecipients;
            }
            finally
            {
                ComInteropScope.TryRelease(
                    recipients,
                    LogCategories.FileLink,
                    "Failed to release password Recipients COM object.");
            }
        }

        private static void AddResolvedRecipients(
            Outlook.Recipients recipients,
            List<string> addresses,
            Outlook.OlMailRecipientType type,
            string composeKey,
            List<string> resolvedRecipients)
        {
            if (recipients == null || addresses == null || addresses.Count == 0)
            {
                return;
            }
            for (int i = 0; i < addresses.Count; i++)
            {
                string address = addresses[i] ?? string.Empty;
                Outlook.Recipient recipient = null;
                try
                {
                    recipient = recipients.Add(address);
                    if (recipient == null)
                    {
                        throw new InvalidOperationException("Recipient could not be added.");
                    }

                    recipient.Type = (int)type;
                    bool resolved = recipient.Resolve();
                    if (!resolved)
                    {
                        throw new InvalidOperationException(
                            "Recipient could not be resolved (composeKey="
                            + (composeKey ?? string.Empty)
                            + ", address="
                            + address
                            + ", type="
                            + type.ToString()
                            + ").");
                    }

                    RecipientAddressList.AddUniqueRecipient(resolvedRecipients, address);
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        recipient,
                        LogCategories.FileLink,
                        "Failed to release password Recipient COM object.");
                }
            }
        }

        private static bool IsDispatchUsable(SeparatePasswordDispatchEntry dispatch)
        {
            if (dispatch == null || string.IsNullOrWhiteSpace(dispatch.Password))
            {
                return false;
            }

            return dispatch.IsPlainText
                ? !string.IsNullOrWhiteSpace(dispatch.PlainText)
                : !string.IsNullOrWhiteSpace(dispatch.Html);
        }

        private static void ApplySeparatePasswordBody(Outlook.MailItem mail, SeparatePasswordDispatchEntry dispatch)
        {
            if (mail == null || dispatch == null)
            {
                return;
            }

            if (dispatch.IsPlainText)
            {
                mail.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
                mail.Body = PlainTextUtilities.NormalizeCrLfAndTrim(dispatch.PlainText);
                return;
            }

            mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            mail.HTMLBody = dispatch.Html ?? string.Empty;
        }

        private string ApplyAndVerifySeparatePasswordSender(
            Outlook.MailItem mail,
            SeparatePasswordDispatchEntry dispatch,
            string composeKey,
            bool requireSourceIdentity)
        {
            if (mail == null || dispatch == null)
            {
                if (requireSourceIdentity)
                {
                    throw new InvalidOperationException("Separate password sender identity is unavailable.");
                }
                return string.Empty;
            }

            string expectedSenderEmail = NormalizeSmtpAddress(dispatch.SenderEmail);
            string requestedAccountSmtp = NormalizeSmtpAddress(dispatch.SendUsingAccountSmtpAddress);
            string requestedSentOnBehalf = (dispatch.SentOnBehalfOfName ?? string.Empty).Trim();
            bool accountApplied = string.IsNullOrWhiteSpace(requestedAccountSmtp)
                || TrySetSeparatePasswordSendUsingAccount(mail, requestedAccountSmtp, composeKey);
            bool sentOnBehalfApplied = string.IsNullOrWhiteSpace(requestedSentOnBehalf)
                || TrySetSeparatePasswordSentOnBehalf(mail, requestedSentOnBehalf, composeKey);
            string effectiveSenderEmail = ResolveSeparatePasswordEffectiveSenderEmail(mail, composeKey);
            bool senderMatches = !string.IsNullOrWhiteSpace(expectedSenderEmail)
                                 && string.Equals(
                                     effectiveSenderEmail,
                                     expectedSenderEmail,
                                     StringComparison.OrdinalIgnoreCase);

            NextcloudTalkAddIn.LogFileLinkMessage(
                "Separate password sender prepared (composeKey="
                + (composeKey ?? string.Empty)
                + ", requireSourceIdentity="
                + requireSourceIdentity.ToString(CultureInfo.InvariantCulture)
                + ", requestedAccount="
                + (!string.IsNullOrWhiteSpace(requestedAccountSmtp)).ToString(CultureInfo.InvariantCulture)
                + ", accountApplied="
                + accountApplied.ToString(CultureInfo.InvariantCulture)
                + ", requestedSentOnBehalf="
                + (!string.IsNullOrWhiteSpace(requestedSentOnBehalf)).ToString(CultureInfo.InvariantCulture)
                + ", sentOnBehalfApplied="
                + sentOnBehalfApplied.ToString(CultureInfo.InvariantCulture)
                + ", effectiveSenderResolved="
                + (!string.IsNullOrWhiteSpace(effectiveSenderEmail)).ToString(CultureInfo.InvariantCulture)
                + ", sourceSenderMatches="
                + senderMatches.ToString(CultureInfo.InvariantCulture)
                + ").");

            if (!requireSourceIdentity)
            {
                return effectiveSenderEmail;
            }
            if (string.IsNullOrWhiteSpace(expectedSenderEmail))
            {
                throw new InvalidOperationException("Separate password source sender identity could not be resolved.");
            }
            if (!accountApplied)
            {
                throw new InvalidOperationException("Separate password sender account could not be applied and verified.");
            }
            if (!sentOnBehalfApplied)
            {
                throw new InvalidOperationException("Separate password sent-on-behalf identity could not be applied and verified.");
            }
            if (!senderMatches)
            {
                throw new InvalidOperationException("Separate password effective sender does not match the source message sender.");
            }

            return effectiveSenderEmail;
        }

        private bool TrySetSeparatePasswordSentOnBehalf(Outlook.MailItem mail, string sentOnBehalfOfName, string composeKey)
        {
            if (mail == null || string.IsNullOrWhiteSpace(sentOnBehalfOfName))
            {
                return false;
            }

            try
            {
                mail.SentOnBehalfOfName = sentOnBehalfOfName.Trim();
                string appliedValue = (mail.SentOnBehalfOfName ?? string.Empty).Trim();
                bool applied = !string.IsNullOrWhiteSpace(appliedValue);
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password sent-on-behalf identity applied (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", verified="
                    + applied.ToString(CultureInfo.InvariantCulture)
                    + ").");
                return applied;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Failed to set separate password sent-on-behalf identity (composeKey=" + (composeKey ?? string.Empty) + ").",
                    ex);
                return false;
            }
        }

        private bool TrySetSeparatePasswordSendUsingAccount(Outlook.MailItem mail, string smtpAddress, string composeKey)
        {
            if (mail == null || _owner.OutlookApplication == null || string.IsNullOrWhiteSpace(smtpAddress))
            {
                return false;
            }

            Outlook.NameSpace session = null;
            Outlook.Accounts accounts = null;
            try
            {
                session = _owner.OutlookApplication.Session;
                if (session == null)
                {
                    return false;
                }

                accounts = session.Accounts;
                if (accounts == null)
                {
                    return false;
                }

                int count = accounts.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Account account = null;
                    try
                    {
                        account = accounts[i];
                        string accountSmtp = EmailSignaturePolicyService.NormalizeEmail(account != null ? account.SmtpAddress : string.Empty);
                        if (!string.Equals(accountSmtp, smtpAddress, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        mail.SendUsingAccount = account;
                        string appliedAccountSmtp = NormalizeSmtpAddress(
                            OutlookRecipientResolverController.ResolveSendUsingAccountSmtpAddress(
                                mail,
                                LogCategories.FileLink,
                                "separate password",
                                " (composeKey=" + (composeKey ?? string.Empty) + ")"));
                        bool verified = string.Equals(
                            appliedAccountSmtp,
                            smtpAddress,
                            StringComparison.OrdinalIgnoreCase);
                        NextcloudTalkAddIn.LogFileLinkMessage(
                            "Separate password send account applied (composeKey="
                            + (composeKey ?? string.Empty)
                            + ", verified="
                            + verified.ToString(CultureInfo.InvariantCulture)
                            + ").");
                        return verified;
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(
                            account,
                            LogCategories.FileLink,
                            "Failed to release separate password Account COM object.");
                    }
                }

                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password send account not found (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", hasRequestedAccount=True).");
                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Failed to apply separate password send account (composeKey=" + (composeKey ?? string.Empty) + ").",
                    ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(accounts, LogCategories.FileLink, "Failed to release separate password Accounts COM object.");
                ComInteropScope.TryRelease(session, LogCategories.FileLink, "Failed to release separate password Session COM object.");
            }
        }

        private string ResolveSeparatePasswordEffectiveSenderEmail(Outlook.MailItem mail, string composeKey)
        {
            string resolved = OutlookRecipientResolverController.ResolveEffectiveSenderSmtpAddress(
                mail,
                _owner.OutlookApplication,
                LogCategories.FileLink,
                "separate password",
                " (composeKey=" + (composeKey ?? string.Empty) + ")",
                false);
            return NormalizeSmtpAddress(resolved);
        }

        private static SeparatePasswordSignatureSnapshot
            BuildSeparatePasswordSignatureSnapshot(
                BackendPolicyStatus policyStatus,
                AddinSettings settings,
                string composeKey)
        {
            var snapshot = new SeparatePasswordSignatureSnapshot
            {
                Active = false,
                Reason = "settings_incomplete",
                UserEmail = string.Empty,
                Html = string.Empty,
                PlainText = string.Empty
            };

            if (settings == null
                || string.IsNullOrWhiteSpace(settings.ServerUrl)
                || string.IsNullOrWhiteSpace(settings.Username)
                || string.IsNullOrWhiteSpace(settings.AppPassword))
            {
                LogSeparatePasswordSignatureSnapshot(composeKey, snapshot);
                return snapshot;
            }

            if (policyStatus == null
                || !policyStatus.FetchSucceeded)
            {
                snapshot.Reason = "policy_snapshot_unavailable";
                LogSeparatePasswordSignatureSnapshot(composeKey, snapshot);
                return snapshot;
            }
            var policy = new EmailSignaturePolicyService(policyStatus, settings).Resolve();
            if (!policy.Active)
            {
                snapshot.Reason = policy.Reason;
                LogSeparatePasswordSignatureSnapshot(composeKey, snapshot);
                return snapshot;
            }

            string sanitized;
            try
            {
                sanitized = HtmlTemplateSanitizer.SanitizeEmailSignatureTemplateHtml(policy.TemplateHtml);
            }
            catch (InvalidOperationException ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Separate password signature snapshot sanitization failed (composeKey=" + (composeKey ?? string.Empty) + ").",
                    ex);
                snapshot.Reason = "sanitization_failed";
                LogSeparatePasswordSignatureSnapshot(composeKey, snapshot);
                return snapshot;
            }
            if (string.IsNullOrWhiteSpace(sanitized))
            {
                snapshot.Reason = "sanitized_empty";
                LogSeparatePasswordSignatureSnapshot(composeKey, snapshot);
                return snapshot;
            }

            snapshot.UserEmail = NormalizeSmtpAddress(policy.UserEmail);
            if (string.IsNullOrWhiteSpace(snapshot.UserEmail))
            {
                snapshot.Reason = "signature_user_email_invalid";
                LogSeparatePasswordSignatureSnapshot(composeKey, snapshot);
                return snapshot;
            }

            snapshot.Active = true;
            snapshot.Reason = "active";
            snapshot.Html = sanitized;
            snapshot.PlainText = EmailSignatureContentBuilder.BuildPlainText(sanitized);
            LogSeparatePasswordSignatureSnapshot(composeKey, snapshot);
            return snapshot;
        }

        private void ApplySeparatePasswordBackendSignature(
            Outlook.MailItem mail,
            SeparatePasswordDispatchEntry dispatch,
            string effectiveSenderEmail,
            SeparatePasswordSignatureSnapshot signatureSnapshot,
            string composeKey)
        {
            if (mail == null || dispatch == null || signatureSnapshot == null || !signatureSnapshot.Active)
            {
                return;
            }

            string normalizedEffectiveSender = NormalizeSmtpAddress(effectiveSenderEmail);
            if (!string.Equals(
                normalizedEffectiveSender,
                signatureSnapshot.UserEmail,
                StringComparison.OrdinalIgnoreCase))
            {
                LogSeparatePasswordSignatureSkipped(composeKey, "effective_identity_mismatch");
                return;
            }

            if (dispatch.IsPlainText)
            {
                string plainText = signatureSnapshot.PlainText;
                if (string.IsNullOrWhiteSpace(plainText))
                {
                    LogSeparatePasswordSignatureSkipped(composeKey, "plain_text_empty");
                    return;
                }

                mail.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
                mail.Body = CombinePlainTextSegments(mail.Body, plainText);
            }
            else
            {
                mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                mail.HTMLBody = AppendHtmlSignature(mail.HTMLBody, signatureSnapshot.Html);
            }

            NextcloudTalkAddIn.LogFileLinkMessage(
                "Separate password backend signature applied (composeKey="
                + (composeKey ?? string.Empty)
                + ", plainText="
                + dispatch.IsPlainText.ToString(CultureInfo.InvariantCulture)
                + ").");
        }

        private void
            ApplySeparatePasswordBackendSignatureToDisplayedFallback(
                Outlook.MailItem mail,
                SeparatePasswordDispatchEntry dispatch,
                SeparatePasswordSignatureSnapshot signatureSnapshot,
                string composeKey)
        {
            if (mail == null
                || dispatch == null
                || signatureSnapshot == null
                || !signatureSnapshot.Active)
            {
                return;
            }

            string effectiveSenderEmail =
                ResolveSeparatePasswordEffectiveSenderEmail(
                    mail,
                    composeKey);
            if (!string.Equals(
                    effectiveSenderEmail,
                    signatureSnapshot.UserEmail,
                    StringComparison.OrdinalIgnoreCase))
            {
                LogSeparatePasswordSignatureSkipped(
                    composeKey,
                    "fallback_effective_identity_mismatch");
                return;
            }

            string signatureContent = dispatch.IsPlainText
                ? signatureSnapshot.PlainText
                : EmailSignatureContentBuilder.BuildManagedHtml(
                    signatureSnapshot.Html);
            if (string.IsNullOrWhiteSpace(signatureContent))
            {
                LogSeparatePasswordSignatureSkipped(
                    composeKey,
                    dispatch.IsPlainText
                        ? "fallback_plain_text_empty"
                        : "fallback_html_empty");
                return;
            }

            try
            {
                ManagedEmailSignatureController.EmailSignatureReconcileResult result =
                    _passwordManagedEmailSignatureController
                        .ApplyManagedEmailSignature(
                            mail,
                            false,
                            dispatch.IsPlainText,
                            signatureContent,
                            false,
                            composeKey,
                            "separate_password_fallback");
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password fallback signature reconciled after display (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", success="
                    + (result != null && result.Success)
                        .ToString(CultureInfo.InvariantCulture)
                    + ", source="
                    + (result != null
                        ? (result.Source ?? "n/a")
                        : "missing")
                    + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Separate password fallback signature reconciliation failed (composeKey="
                    + (composeKey ?? string.Empty)
                    + ").",
                    ex);
            }
        }

        private static string NormalizeSmtpAddress(string value)
        {
            string normalized = EmailSignaturePolicyService.NormalizeEmail(value);
            return normalized.IndexOf("@", StringComparison.Ordinal) > 0 ? normalized : string.Empty;
        }

        private static string CombinePlainTextSegments(string body, string signature)
        {
            string normalizedBody = PlainTextUtilities.NormalizeCrLfAndTrim(body);
            string normalizedSignature = PlainTextUtilities.NormalizeCrLfAndTrim(signature);
            if (string.IsNullOrWhiteSpace(normalizedBody))
            {
                return normalizedSignature;
            }
            if (string.IsNullOrWhiteSpace(normalizedSignature))
            {
                return normalizedBody;
            }
            return normalizedBody + "\r\n\r\n" + normalizedSignature;
        }

        private static string AppendHtmlSignature(string html, string sanitizedSignature)
        {
            string existing = html ?? string.Empty;
            string signatureBlock = "<br><br>" + EmailSignatureContentBuilder.BuildManagedHtml(sanitizedSignature);
            int bodyEnd = existing.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
            if (bodyEnd >= 0)
            {
                return existing.Insert(bodyEnd, signatureBlock);
            }
            return existing + signatureBlock;
        }

        private static void LogSeparatePasswordSignatureSkipped(string composeKey, string reason)
        {
            NextcloudTalkAddIn.LogFileLinkMessage(
                "Separate password backend signature skipped (composeKey="
                + (composeKey ?? string.Empty)
                + ", reason="
                + (reason ?? "n/a")
                + ").");
        }

        private static void LogSeparatePasswordSignatureSnapshot(
            string composeKey,
            SeparatePasswordSignatureSnapshot snapshot)
        {
            NextcloudTalkAddIn.LogFileLinkMessage(
                "Separate password backend signature snapshot prepared (composeKey="
                + (composeKey ?? string.Empty)
                + ", active="
                + (snapshot != null && snapshot.Active).ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (snapshot != null ? (snapshot.Reason ?? "n/a") : "missing")
                + ", hasUserEmail="
                + (snapshot != null && !string.IsNullOrWhiteSpace(snapshot.UserEmail)).ToString(CultureInfo.InvariantCulture)
                + ", hasHtml="
                + (snapshot != null && !string.IsNullOrWhiteSpace(snapshot.Html)).ToString(CultureInfo.InvariantCulture)
                + ", hasPlainText="
                + (snapshot != null && !string.IsNullOrWhiteSpace(snapshot.PlainText)).ToString(CultureInfo.InvariantCulture)
                + ").");
        }

        private static string BuildSeparatePasswordMailSubject(SeparatePasswordDispatchEntry dispatch)
        {
            string baseSubject = Strings.SharingPasswordMailSubject;
            string shareLabel = dispatch != null ? (dispatch.ShareLabel ?? string.Empty).Trim() : string.Empty;
            if (string.IsNullOrWhiteSpace(shareLabel))
            {
                return baseSubject;
            }
            return string.Format(CultureInfo.CurrentCulture, Strings.SharingPasswordMailSubjectWithLabel, shareLabel);
        }

    }
}
