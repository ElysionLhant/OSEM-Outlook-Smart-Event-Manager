using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace OSEMAddIn.Services
{
    internal static class DebugLogger
    {
        private static readonly object Gate = new object();
        private static readonly string LogPath = BuildLogPath();

        public static void Log(string message)
        {
            try
            {
                var line = $"{DateTime.Now.ToString("O", CultureInfo.InvariantCulture)}\t{message}";
                lock (Gate)
                {
                    File.AppendAllText(LogPath, line + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch
            {
                // Swallow logging failures; diagnostics should not break runtime behavior.
            }
        }

        private static string BuildLogPath()
        {
            var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var folder = Path.Combine(root, "OSEM");
            Directory.CreateDirectory(folder);
            return Path.Combine(folder, "addin-debug.log");
        }
    }
}
