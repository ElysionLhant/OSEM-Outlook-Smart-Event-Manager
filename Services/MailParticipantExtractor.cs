#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OSEMAddIn.Services
{
    internal static class MailParticipantExtractor
    {
    public static string[] Capture(Outlook.MailItem mailItem)
        {
            if (mailItem is null)
            {
                return Array.Empty<string>();
            }

            var participants = new HashSet<string>(StringComparer.Ordinal);

            void Add(string? value)
            {
                var normalized = Normalize(value);
                if (!string.IsNullOrEmpty(normalized))
                {
                    participants.Add(normalized);
                }
            }

            try
            {
                Add(mailItem.SenderEmailAddress);
            }
            catch (COMException)
            {
                // ignore sender address failures
            }

            Outlook.AddressEntry? senderEntry = null;
            try
            {
                senderEntry = mailItem.Sender;
                if (senderEntry is not null)
                {
                    Add(senderEntry.Address);
                    Add(senderEntry.Name);

                    Outlook.ExchangeUser? exchangeUser = null;
                    try
                    {
                        exchangeUser = senderEntry.GetExchangeUser();
                        if (exchangeUser is not null)
                        {
                            Add(exchangeUser.PrimarySmtpAddress);
                        }
                    }
                    catch (COMException)
                    {
                        // ignore exchange resolution failures
                    }
                    finally
                    {
                        if (exchangeUser is not null)
                        {
                            Marshal.ReleaseComObject(exchangeUser);
                        }
                    }

                    Outlook.ExchangeDistributionList? distributionList = null;
                    try
                    {
                        distributionList = senderEntry.GetExchangeDistributionList();
                        if (distributionList is not null)
                        {
                            Add(distributionList.PrimarySmtpAddress);
                        }
                    }
                    catch (COMException)
                    {
                        // ignore exchange distribution failures
                    }
                    finally
                    {
                        if (distributionList is not null)
                        {
                            Marshal.ReleaseComObject(distributionList);
                        }
                    }
                }
            }
            catch (COMException)
            {
                // ignore sender lookup failures
            }
            finally
            {
                if (senderEntry is not null)
                {
                    Marshal.ReleaseComObject(senderEntry);
                }
            }

            Outlook.Recipients? recipients = null;
            try
            {
                recipients = mailItem.Recipients;
                if (recipients is not null)
                {
                    var count = recipients.Count;
                    for (var index = 1; index <= count; index++)
                    {
                        Outlook.Recipient? recipient = null;
                        try
                        {
                            recipient = recipients[index];
                            if (recipient is null)
                            {
                                continue;
                            }

                            Add(recipient.Address);
                            Add(recipient.Name);

                            Outlook.AddressEntry? addressEntry = null;
                            try
                            {
                                addressEntry = recipient.AddressEntry;
                                if (addressEntry is not null)
                                {
                                    Add(addressEntry.Address);
                                    Add(addressEntry.Name);

                                    Outlook.ExchangeUser? exchangeUser = null;
                                    try
                                    {
                                        exchangeUser = addressEntry.GetExchangeUser();
                                        if (exchangeUser is not null)
                                        {
                                            Add(exchangeUser.PrimarySmtpAddress);
                                        }
                                    }
                                    catch (COMException)
                                    {
                                        // ignore failures
                                    }
                                    finally
                                    {
                                        if (exchangeUser is not null)
                                        {
                                            Marshal.ReleaseComObject(exchangeUser);
                                        }
                                    }

                                    Outlook.ExchangeDistributionList? distributionList = null;
                                    try
                                    {
                                        distributionList = addressEntry.GetExchangeDistributionList();
                                        if (distributionList is not null)
                                        {
                                            Add(distributionList.PrimarySmtpAddress);
                                        }
                                    }
                                    catch (COMException)
                                    {
                                        // ignore failures
                                    }
                                    finally
                                    {
                                        if (distributionList is not null)
                                        {
                                            Marshal.ReleaseComObject(distributionList);
                                        }
                                    }
                                }
                            }
                            catch (COMException)
                            {
                                // ignore address entry failures
                            }
                            finally
                            {
                                if (addressEntry is not null)
                                {
                                    Marshal.ReleaseComObject(addressEntry);
                                }
                            }
                        }
                        catch (COMException)
                        {
                            // ignore recipient failures
                        }
                        finally
                        {
                            if (recipient is not null)
                            {
                                Marshal.ReleaseComObject(recipient);
                            }
                        }
                    }
                }
            }
            catch (COMException)
            {
                // ignore recipients collection failures
            }
            finally
            {
                if (recipients is not null)
                {
                    Marshal.ReleaseComObject(recipients);
                }
            }

            return participants.Count == 0 ? Array.Empty<string>() : participants.ToArray();
        }

        public static string[] Normalize(IEnumerable<string>? participants)
        {
            if (participants is null)
            {
                return Array.Empty<string>();
            }

            var set = new HashSet<string>(StringComparer.Ordinal);
            foreach (var participant in participants)
            {
                var normalized = Normalize(participant);
                if (!string.IsNullOrEmpty(normalized))
                {
                    set.Add(normalized);
                }
            }

            return set.Count == 0 ? Array.Empty<string>() : set.ToArray();
        }

        public static bool Intersects(IEnumerable<string>? candidateParticipants, ISet<string>? knownParticipants)
        {
            if (candidateParticipants is null || knownParticipants is null || knownParticipants.Count == 0)
            {
                return false;
            }

            foreach (var participant in candidateParticipants)
            {
                if (string.IsNullOrWhiteSpace(participant))
                {
                    continue;
                }

                if (knownParticipants.Contains(participant))
                {
                    return true;
                }
            }

            return false;
        }

        public static double ComputeOverlapScore(IEnumerable<string>? first, IEnumerable<string>? second)
        {
            var firstSet = Normalize(first);
            var secondSet = Normalize(second);

            if (firstSet.Length == 0 || secondSet.Length == 0)
            {
                return 0d;
            }

            var secondLookup = new HashSet<string>(secondSet, StringComparer.Ordinal);
            var shared = 0;

            foreach (var participant in firstSet)
            {
                if (secondLookup.Contains(participant))
                {
                    shared++;
                }
            }

            var denominator = Math.Max(firstSet.Length, secondSet.Length);
            return denominator == 0 ? 0d : (double)shared / denominator;
        }

        internal static string Normalize(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = value!.Trim().Trim('"', '\'', '<', '>', ';');
            if (normalized.Length == 0)
            {
                return string.Empty;
            }

            if (normalized.StartsWith("SMTP:", StringComparison.OrdinalIgnoreCase) ||
                normalized.StartsWith("EX:", StringComparison.OrdinalIgnoreCase))
            {
                var colonIndex = normalized.IndexOf(':');
                if (colonIndex >= 0 && colonIndex + 1 < normalized.Length)
                {
                    normalized = normalized.Substring(colonIndex + 1).Trim();
                }
            }

            if (normalized.StartsWith("MAILTO:", StringComparison.OrdinalIgnoreCase))
            {
                normalized = normalized.Substring(7).Trim();
            }

            return normalized.Length == 0 ? string.Empty : normalized.ToUpperInvariant();
        }
    }
}
