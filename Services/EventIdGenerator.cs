using System;

namespace OSEMAddIn.Services
{
    internal static class EventIdGenerator
    {
        public static string Generate(string prefix = "EVT")
        {
            var timestamp = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss");
            var random = Guid.NewGuid().ToString("N").Substring(0, 6).ToUpperInvariant();
            return $"{prefix}-{timestamp}-{random}";
        }
    }
}
