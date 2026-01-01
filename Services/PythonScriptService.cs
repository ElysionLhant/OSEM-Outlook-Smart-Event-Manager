using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal sealed class PythonScriptService
    {
        private readonly string _scriptRoot;
        private readonly string _metaPath;
        private Dictionary<string, ScriptMetadata> _metadataCache = new();

        private class ScriptMetadata
        {
            public string Description { get; set; } = "";
            public bool IsGlobal { get; set; }
            public List<string> AssociatedTemplateIds { get; set; } = new();
        }

        public PythonScriptService(string? scriptRoot = null)
        {
            _scriptRoot = scriptRoot ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Scripts");
            Directory.CreateDirectory(_scriptRoot);
            _metaPath = Path.Combine(_scriptRoot, "scripts_meta.json");
            LoadMetadata();
        }

        private void LoadMetadata()
        {
            if (File.Exists(_metaPath))
            {
                try
                {
                    var json = File.ReadAllText(_metaPath);
                    _metadataCache = JsonConvert.DeserializeObject<Dictionary<string, ScriptMetadata>>(json) 
                                     ?? new Dictionary<string, ScriptMetadata>();
                }
                catch
                {
                    _metadataCache = new Dictionary<string, ScriptMetadata>();
                }
            }
        }

        private void SaveMetadata()
        {
            try
            {
                var json = JsonConvert.SerializeObject(_metadataCache, Formatting.Indented);
                File.WriteAllText(_metaPath, json);
            }
            catch
            {
                // ignore save errors
            }
        }

        public void UpdateScriptMetadata(PythonScriptDefinition script)
        {
            if (!_metadataCache.ContainsKey(script.ScriptId))
            {
                _metadataCache[script.ScriptId] = new ScriptMetadata();
            }
            
            var meta = _metadataCache[script.ScriptId];
            meta.Description = script.Description;
            meta.IsGlobal = script.IsGlobal;
            meta.AssociatedTemplateIds = script.AssociatedTemplateIds;
            
            SaveMetadata();
        }

        public void SaveScript(string fileName, string content)
        {
            var path = Path.Combine(_scriptRoot, fileName);
            File.WriteAllText(path, content);
        }

        public IReadOnlyList<PythonScriptDefinition> DiscoverScripts()
        {
            var scripts = Directory.EnumerateFiles(_scriptRoot, "*.py")
                .Select(path => {
                    var id = Path.GetFileNameWithoutExtension(path);
                    var meta = _metadataCache.ContainsKey(id) ? _metadataCache[id] : new ScriptMetadata();
                    
                    return new PythonScriptDefinition
                    {
                        ScriptId = id,
                        DisplayName = Path.GetFileName(path),
                        ScriptPath = path,
                        Description = string.IsNullOrEmpty(meta.Description) ? Properties.Resources.External_Python_Analysis_Script : meta.Description,
                        IsGlobal = meta.IsGlobal,
                        AssociatedTemplateIds = meta.AssociatedTemplateIds ?? new List<string>()
                    };
                })
                .ToList();

            if (scripts.Count == 0)
            {
                // No default scripts
            }

            return scripts;
        }

        public Task<int> ExecuteAsync(PythonScriptDefinition script, PythonScriptExecutionContext context)
        {
            if (!File.Exists(script.ScriptPath))
            {
                throw new FileNotFoundException(Properties.Resources.Script_not_found, script.ScriptPath);
            }

            var contextFile = WriteContextToTempFile(context);

            var processStartInfo = new ProcessStartInfo
            {
                FileName = "python.exe",
                Arguments = $"\"{script.ScriptPath}\" \"{contextFile}\"",
                WorkingDirectory = Path.GetDirectoryName(script.ScriptPath) ?? _scriptRoot,
                CreateNoWindow = false,
                UseShellExecute = true
            };

            var process = Process.Start(processStartInfo);
            if (process is null)
            {
                TryDeleteContextFile(contextFile);
                throw new InvalidOperationException(Properties.Resources.Unable_to_start_Python_process);
            }

            // .NET Framework 4.8 doesn't have Process.WaitForExitAsync; use synchronous wait and return a completed task
            process.WaitForExit();
            TryDeleteContextFile(contextFile);
            return Task.FromResult(process.ExitCode);
        }

        private static string WriteContextToTempFile(PythonScriptExecutionContext context)
        {
            var filePath = Path.Combine(Path.GetTempPath(), $"osem-script-context-{Guid.NewGuid():N}.json");
            var json = JsonConvert.SerializeObject(context, Formatting.Indented);
            File.WriteAllText(filePath, json);
            return filePath;
        }

        private static void TryDeleteContextFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // ignore cleanup errors
            }
        }
    }
}
