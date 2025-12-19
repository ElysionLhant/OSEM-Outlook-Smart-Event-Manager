using System;
using System.Text;

namespace OSEMAddIn.Services
{
    public static class EncodingRepair
    {
        // Common legacy encodings that might be misinterpreted as UTF-8:
        // 936 (GBK), 54936 (GB18030), 932 (Shift-JIS), 949 (EUC-KR), 950 (Big5)
        private static readonly int[] CandidateCodePages = new[] { 936, 54936, 932, 949, 950 };

        /// <summary>
        /// Attempts to repair a mojibake string by reversing common encoding misinterpretations.
        /// </summary>
        /// <param name="input">The potentially garbled string.</param>
        /// <param name="validator">A predicate to validate if the repaired string is correct (e.g. matches a regex or starts with a prefix).</param>
        /// <param name="repaired">The repaired string if successful.</param>
        /// <returns>True if a repair was found and validated; otherwise false.</returns>
        public static bool TryFix(string input, Func<string, bool> validator, out string repaired)
        {
            repaired = string.Empty;
            if (string.IsNullOrEmpty(input))
            {
                return false;
            }

            var utf8 = Encoding.UTF8;

            foreach (var codePage in CandidateCodePages)
            {
                try
                {
                    var encoding = Encoding.GetEncoding(codePage);
                    // Reverse the mojibake: 
                    // Assumption: The 'input' is what happens when 'OriginalBytes' (which were 'encoding') 
                    // were incorrectly read as UTF-8.
                    // So we get the bytes of 'input' using 'encoding' (which is actually wrong, but...)
                    // Wait, the logic is:
                    // Original: GBK Bytes -> Read as UTF-8 -> Garbage String (Input)
                    // To Fix: Garbage String (Input) -> Get Bytes as UTF-8? NO.
                    // If I have Garbage, and I want to get back the GBK bytes.
                    // Garbage characters are formed by taking GBK bytes and mapping them to Unicode via UTF-8 (or Latin1).
                    // If it was Latin1 (ISO-8859-1), we would do: Input -> GetBytes(Latin1) -> Bytes -> GetString(GBK).
                    
                    // BUT, in the case of "涓婚" (Subject):
                    // Python: '主题'.encode('utf-8').decode('gb18030') -> '涓婚'
                    // This means: Original was UTF-8 bytes. It was interpreted as GB18030.
                    // So we have '涓婚'. We want to get back '主题'.
                    // We need to: '涓婚'.encode('gb18030') -> Bytes (which are actually UTF-8) -> decode('utf-8').
                    
                    // So: Input.GetBytes(CodePage) -> Bytes -> UTF8.GetString(Bytes).
                    
                    var bytes = encoding.GetBytes(input);
                    var candidate = utf8.GetString(bytes);

                    // Only accept if the string actually changed (avoid false positives on plain ASCII)
                    if (!string.Equals(input, candidate, StringComparison.Ordinal))
                    {
                        if (validator(candidate))
                        {
                            repaired = candidate;
                            return true;
                        }
                    }
                }
                catch
                {
                    // Ignore encoding errors
                }
            }

            return false;
        }
    }
}
