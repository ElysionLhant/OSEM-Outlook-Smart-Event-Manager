using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal sealed class RegexExtractionService
    {
        public Dictionary<string, string> Extract(MailItem mailItem, DashboardTemplate template)
        {
            var results = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (mailItem is null || template is null)
            {
                return results;
            }

            var body = mailItem.Body ?? string.Empty;

            if (template.FieldRegexes != null)
            {
                foreach (var kvp in template.FieldRegexes)
                {
                    var key = kvp.Key;
                    var pattern = kvp.Value;

                    if (string.IsNullOrWhiteSpace(pattern)) continue;

                    try
                    {
                        var regex = new Regex(pattern, RegexOptions.IgnoreCase);
                        TryCapture(regex, body, key, results);
                    }
                    catch
                    {
                        // Ignore invalid regex
                    }
                }
            }

            return results;
        }

        private static void TryCapture(Regex regex, string input, string key, IDictionary<string, string> result)
        {
            var match = regex.Match(input);
            if (!match.Success)
            {
                return;
            }

            var value = match.Groups["value"].Value;
            if (!string.IsNullOrWhiteSpace(value))
            {
                result[key] = value.Trim();
            }
            else if (match.Groups.Count > 1)
            {
                // Fallback to first group if "value" group is not present
                result[key] = match.Groups[1].Value.Trim();
            }
        }
    }
}
