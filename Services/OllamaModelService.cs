using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace OSEMAddIn.Services
{
    internal sealed class OllamaModelService
    {
        private readonly string[] _fallbackModels = { "llama2", "mistral", "gemma" };

        public async Task<IReadOnlyList<string>> GetModelsAsync()
        {
            try
            {
                return await Task.Run(QueryModels).ConfigureAwait(false);
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"OllamaModelService.GetModelsAsync falling back to defaults: {ex}");
                return _fallbackModels;
            }
        }

        private IReadOnlyList<string> QueryModels()
        {
            var executablePath = ResolveOllamaExecutablePath();
            var startInfo = new ProcessStartInfo
            {
                FileName = executablePath ?? "ollama",
                Arguments = "list",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) && executablePath is not null)
            {
                var directory = Path.GetDirectoryName(executablePath);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    startInfo.WorkingDirectory = directory;
                }
            }

            Process? process;
            try
            {
                process = Process.Start(startInfo);
            }
            catch (System.ComponentModel.Win32Exception)
            {
                return _fallbackModels;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"OllamaModelService failed to start process: {ex}");
                return _fallbackModels;
            }

            using (process)
            {
                if (process is null)
                {
                    DebugLogger.Log($"OllamaModelService failed to start process for '{startInfo.FileName}'");
                    return _fallbackModels;
                }

                if (!process.WaitForExit(10000))
                {
                    try
                    {
                        process.Kill();
                    }
                    catch
                    {
                        // ignored - best effort to stop runaway process.
                    }

                    DebugLogger.Log($"OllamaModelService timed out waiting for '{startInfo.FileName} {startInfo.Arguments}'");
                    return _fallbackModels;
                }

                var output = process.StandardOutput.ReadToEnd();
                var errors = process.StandardError.ReadToEnd();

                if (process.ExitCode != 0)
                {
                    DebugLogger.Log($"OllamaModelService detected exit code {process.ExitCode} from '{startInfo.FileName}': {errors.Trim()}");
                    return _fallbackModels;
                }

                if (!string.IsNullOrWhiteSpace(errors))
                {
                    DebugLogger.Log($"OllamaModelService stderr from '{startInfo.FileName}': {errors.Trim()}");
                }

                if (string.IsNullOrWhiteSpace(output))
                {
                    DebugLogger.Log("OllamaModelService received empty output when querying models.");
                    return _fallbackModels;
                }

                var models = new List<string>();
                using var reader = new StringReader(output);
                string? line;
                var isFirstLine = true;
                while ((line = reader.ReadLine()) is not null)
                {
                    if (isFirstLine)
                    {
                        // skip header line: NAME ID SIZE MODIFIED
                        isFirstLine = false;
                        continue;
                    }

                    var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length > 0)
                    {
                        models.Add(parts[0]);
                    }
                }

                return models.Count == 0 ? _fallbackModels : models.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            }
        }

        private static string? ResolveOllamaExecutablePath()
        {
            var candidates = new List<string?>
            {
                Environment.GetEnvironmentVariable("OLLAMA_EXE"),
                Environment.GetEnvironmentVariable("OLLAMA_PATH"),
                Environment.GetEnvironmentVariable("OLLAMA_HOME")
            };

            foreach (var candidate in candidates)
            {
                var path = NormalizeExecutableCandidate(candidate);
                if (path is not null)
                {
                    return path;
                }
            }

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                var localAppPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Ollama", "ollama.exe");
                if (File.Exists(localAppPath))
                {
                    return localAppPath;
                }

                var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                if (!string.IsNullOrWhiteSpace(programFiles))
                {
                    var programFilesPath = Path.Combine(programFiles, "Ollama", "ollama.exe");
                    if (File.Exists(programFilesPath))
                    {
                        return programFilesPath;
                    }
                }

                var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
                if (!string.IsNullOrWhiteSpace(programFilesX86))
                {
                    var programFilesX86Path = Path.Combine(programFilesX86, "Ollama", "ollama.exe");
                    if (File.Exists(programFilesX86Path))
                    {
                        return programFilesX86Path;
                    }
                }
            }

            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                const string macDefault = "/Applications/Ollama.app/Contents/MacOS/ollama";
                if (File.Exists(macDefault))
                {
                    return macDefault;
                }
            }

            return null;
        }

        private static string? NormalizeExecutableCandidate(string? candidate)
        {
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return null;
            }

            var trimmed = (candidate ?? string.Empty).Trim();

            if (File.Exists(trimmed))
            {
                return trimmed;
            }

            var combined = Path.Combine(trimmed, RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "ollama.exe" : "ollama");
            return File.Exists(combined) ? combined : null;
        }
    }
}
