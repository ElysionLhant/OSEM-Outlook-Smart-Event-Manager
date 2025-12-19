using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OSEMAddIn.Services
{
    internal static class MailBodyFingerprint
    {
        private const int MaxLength = 512;
    private static readonly Regex HtmlTagRegex = new("<.*?>", RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex WhitespaceRegex = new("\\s+", RegexOptions.Compiled);
    private static readonly Regex QuotedLineRegex = new(@"^\s*>.*$", RegexOptions.Multiline | RegexOptions.Compiled);
    private const int BaselinePrefixLength = 256;

        public static string GetBodyText(Outlook.MailItem mailItem)
        {
            string? body = TryGetBody(mailItem);

            if (string.IsNullOrWhiteSpace(body))
            {
                body = TryGetHtmlAsText(mailItem);
            }

            return body ?? string.Empty;
        }

        public static string Capture(Outlook.MailItem mailItem)
        {
            if (mailItem is null)
            {
                return string.Empty;
            }

            return Capture(GetBodyText(mailItem));
        }

        public static string Capture(string? body)
        {
            if (string.IsNullOrWhiteSpace(body))
            {
                return string.Empty;
            }

            body = QuotedLineRegex.Replace(body!, string.Empty);
            body = WhitespaceRegex.Replace(body, " ").Trim();

            if (body.Length > MaxLength)
            {
                body = body.Substring(0, MaxLength);
            }

            return body.ToUpperInvariant();
        }

        public static bool IsSimilar(string candidate, IReadOnlyCollection<string>? knownFingerprints, double threshold = 0.7)
        {
            if (string.IsNullOrEmpty(candidate) || knownFingerprints is null || knownFingerprints.Count == 0)
            {
                return false;
            }

            foreach (var fingerprint in knownFingerprints)
            {
                if (string.IsNullOrEmpty(fingerprint))
                {
                    continue;
                }

                if (string.Equals(candidate, fingerprint, StringComparison.Ordinal))
                {
                    return true;
                }

                if (ComputeSimilarityScore(candidate, fingerprint) >= threshold)
                {
                    return true;
                }
            }

            return false;
        }

        public static bool MatchesBaseline(string? candidate, string? baseline)
        {
            if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(baseline))
            {
                return false;
            }

            if (string.Equals(candidate, baseline, StringComparison.Ordinal))
            {
                return true;
            }

            var candidateValue = candidate!.Trim();
            var baselineValue = baseline!.Trim();

            if (candidateValue.Length == 0 || baselineValue.Length == 0)
            {
                return false;
            }

            var trimmedCandidate = candidateValue.Length > BaselinePrefixLength
                ? candidateValue.Substring(0, BaselinePrefixLength)
                : candidateValue;

            var trimmedBaseline = baselineValue.Length > BaselinePrefixLength
                ? baselineValue.Substring(0, BaselinePrefixLength)
                : baselineValue;

            if (trimmedCandidate.StartsWith(trimmedBaseline, StringComparison.Ordinal) ||
                trimmedBaseline.StartsWith(trimmedCandidate, StringComparison.Ordinal))
            {
                return true;
            }

            var minLength = Math.Min(trimmedCandidate.Length, trimmedBaseline.Length);
            if (minLength == 0)
            {
                return false;
            }

            var candidateHead = trimmedCandidate.Substring(0, minLength);
            var baselineHead = trimmedBaseline.Substring(0, minLength);

            return candidateHead.Equals(baselineHead, StringComparison.Ordinal);
        }

        public static double ComputeSimilarityScore(string? left, string? right)
        {
            if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right))
            {
                return 0d;
            }

            return ComputeDiceCoefficient(left!, right!);
        }

        private static string? TryGetBody(Outlook.MailItem mailItem)
        {
            try
            {
                return mailItem.Body;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static string? TryGetHtmlAsText(Outlook.MailItem mailItem)
        {
            try
            {
                var html = mailItem.HTMLBody;
                if (string.IsNullOrWhiteSpace(html))
                {
                    return null;
                }

                return HtmlTagRegex.Replace(html, " ");
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static double ComputeDiceCoefficient(string left, string right)
        {
            if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right))
            {
                return 0d;
            }

            var leftBigrams = BuildBigrams(left);
            if (leftBigrams.Count == 0)
            {
                return 0d;
            }

            var rightBigrams = BuildBigrams(right);
            if (rightBigrams.Count == 0)
            {
                return 0d;
            }

            var shared = 0;
            foreach (var bigram in leftBigrams)
            {
                if (rightBigrams.Contains(bigram))
                {
                    shared++;
                }
            }

            return (2d * shared) / (leftBigrams.Count + rightBigrams.Count);
        }

        private static HashSet<string> BuildBigrams(string value)
        {
            var set = new HashSet<string>(StringComparer.Ordinal);
            for (var index = 0; index < value.Length - 1; index++)
            {
                set.Add(value.Substring(index, 2));
            }

            return set;
        }
    }
}
