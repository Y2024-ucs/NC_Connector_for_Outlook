// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    // Owns the short-lived UserProperties and UserProperty COM references used by Talk metadata.
    internal sealed partial class TalkAppointmentController
    {
        internal static bool SetUserProperty(
            Outlook.AppointmentItem appointment,
            string name,
            Outlook.OlUserPropertyType type,
            object value)
        {
            if (appointment == null)
            {
                return false;
            }
            try
            {
                return UseUserProperty(
                    appointment,
                    name,
                    true,
                    type,
                    false,
                    property =>
                    {
                        object currentValue = property.Value;
                        if (!UserPropertyValueEquals(currentValue, value))
                        {
                            property.Value = value;
                        }
                        return true;
                    });
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to set user property '" + name + "'.", ex);
                return false;
            }
        }

        internal static string GetUserPropertyText(Outlook.AppointmentItem appointment, string name)
        {
            if (appointment == null)
            {
                return null;
            }
            try
            {
                return UseUserProperty(
                    appointment,
                    name,
                    false,
                    Outlook.OlUserPropertyType.olText,
                    (string)null,
                    property => property.Value as string);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read user property '" + name + "'.", ex);
                return null;
            }
        }

        internal static bool HasUserProperty(Outlook.AppointmentItem appointment, string name)
        {
            if (appointment == null)
            {
                return false;
            }
            try
            {
                return UseUserProperty(
                    appointment,
                    name,
                    false,
                    Outlook.OlUserPropertyType.olText,
                    false,
                    property => true);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to check user property '" + name + "'.", ex);
                return false;
            }
        }

        internal static bool GetUserPropertyBool(Outlook.AppointmentItem appointment, string name)
        {
            if (appointment == null)
            {
                return false;
            }
            try
            {
                return UseUserProperty(
                    appointment,
                    name,
                    false,
                    Outlook.OlUserPropertyType.olText,
                    false,
                    property =>
                    {
                        object value = property.Value;
                        return ConvertUserPropertyToBool(value);
                    });
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read boolean user property '" + name + "'.", ex);
                return false;
            }
        }

        internal static bool RemoveUserProperty(Outlook.AppointmentItem appointment, string name)
        {
            if (appointment == null)
            {
                return false;
            }
            try
            {
                return UseUserProperty(
                    appointment,
                    name,
                    false,
                    Outlook.OlUserPropertyType.olText,
                    true,
                    property =>
                    {
                        property.Delete();
                        return true;
                    });
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to remove user property '" + name + "'.", ex);
                return false;
            }
        }

        private static TResult UseUserProperty<TResult>(
            Outlook.AppointmentItem appointment,
            string name,
            bool create,
            Outlook.OlUserPropertyType createType,
            TResult missingValue,
            Func<Outlook.UserProperty, TResult> useProperty)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = appointment.UserProperties;
                if (properties == null)
                {
                    return missingValue;
                }

                property = properties[name];
                if (property == null && create)
                {
                    property = properties.Add(name, createType, true, Type.Missing);
                }
                if (property == null)
                {
                    return missingValue;
                }

                return useProperty(property);
            }
            finally
            {
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release Outlook UserProperty COM object.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release Outlook UserProperties COM object.");
            }
        }

        private static bool UserPropertyValueEquals(object currentValue, object newValue)
        {
            if (currentValue == null && newValue == null)
            {
                return true;
            }
            if (currentValue == null || newValue == null)
            {
                return false;
            }

            string currentText = currentValue as string;
            string newText = newValue as string;
            if (currentText != null || newText != null)
            {
                return string.Equals(
                    Convert.ToString(currentValue, CultureInfo.InvariantCulture),
                    Convert.ToString(newValue, CultureInfo.InvariantCulture),
                    StringComparison.Ordinal);
            }

            return object.Equals(currentValue, newValue);
        }

        private static bool ConvertUserPropertyToBool(object value)
        {
            if (value == null)
            {
                return false;
            }
            if (value is bool)
            {
                return (bool)value;
            }
            if (value is int)
            {
                return (int)value != 0;
            }

            string text = value as string;
            if (!string.IsNullOrEmpty(text))
            {
                bool boolParsed;
                if (bool.TryParse(text, out boolParsed))
                {
                    return boolParsed;
                }
                int numericParsed;
                if (int.TryParse(text, out numericParsed))
                {
                    return numericParsed != 0;
                }
            }
            return false;
        }
    }
}
