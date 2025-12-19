using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Newtonsoft.Json;
using OSEMAddIn.Models;
using OSEMAddIn.Views;

namespace OSEMAddIn.Services
{
    internal class TemplatePackageService
    {
        private readonly DashboardTemplateService _templateService;
        private readonly PromptLibraryService _promptService;
        private readonly PythonScriptService _scriptService;

        public TemplatePackageService(
            DashboardTemplateService templateService,
            PromptLibraryService promptService,
            PythonScriptService scriptService)
        {
            _templateService = templateService;
            _promptService = promptService;
            _scriptService = scriptService;
        }

        public void ExportPackage(string filePath, List<DashboardTemplate> templates)
        {
            var package = new TemplatePackage();
            package.Templates = templates;

            foreach (var template in templates)
            {
                // Prompts
                var linkedPrompts = _promptService.GetPrompts()
                    .Where(p => p.TemplateOverrideId != null && 
                                p.TemplateOverrideId.Split(',').Select(id => id.Trim()).Contains(template.TemplateId))
                    .ToList();
                package.Prompts.AddRange(linkedPrompts);

                // Scripts
                var linkedScripts = _scriptService.DiscoverScripts()
                    .Where(s => s.AssociatedTemplateIds.Contains(template.TemplateId))
                    .ToList();
                package.Scripts.AddRange(linkedScripts);

                // Files
                foreach (var attachmentPath in template.AttachmentPaths)
                {
                    if (File.Exists(attachmentPath))
                    {
                        string fileName = Path.GetFileName(attachmentPath);
                        if (!package.Files.ContainsKey(fileName))
                        {
                            package.Files[fileName] = "files/" + fileName;
                        }
                    }
                }
            }

            // Deduplicate
            package.Prompts = package.Prompts.GroupBy(p => p.PromptId).Select(g => g.First()).ToList();
            package.Scripts = package.Scripts.GroupBy(s => s.ScriptId).Select(g => g.First()).ToList();

            if (File.Exists(filePath)) File.Delete(filePath);

            using (var zip = ZipFile.Open(filePath, ZipArchiveMode.Create))
            {
                string json = JsonConvert.SerializeObject(package, Formatting.Indented);
                var entry = zip.CreateEntry("manifest.json");
                using (var writer = new StreamWriter(entry.Open()))
                {
                    writer.Write(json);
                }

                // Write files
                foreach (var kvp in package.Files)
                {
                    string fileName = kvp.Key;
                    string zipPath = kvp.Value;
                    
                    string? fullPath = null;
                    foreach(var t in templates)
                    {
                        var match = t.AttachmentPaths.FirstOrDefault(p => Path.GetFileName(p) == fileName);
                        if (match != null) { fullPath = match; break; }
                    }

                    if (fullPath != null && File.Exists(fullPath))
                    {
                        zip.CreateEntryFromFile(fullPath, zipPath);
                    }
                }
                
                // Write scripts content
                foreach (var script in package.Scripts)
                {
                    if (File.Exists(script.ScriptPath))
                    {
                        string scriptName = Path.GetFileName(script.ScriptPath);
                        zip.CreateEntryFromFile(script.ScriptPath, "scripts/" + scriptName);
                    }
                }
            }
        }

        public void ImportPackage(string filePath)
        {
            using (var zip = ZipFile.OpenRead(filePath))
            {
                var manifestEntry = zip.GetEntry("manifest.json");
                if (manifestEntry == null) throw new Exception("Invalid package: manifest.json missing");

                TemplatePackage? package;
                using (var reader = new StreamReader(manifestEntry.Open()))
                {
                    string json = reader.ReadToEnd();
                    package = JsonConvert.DeserializeObject<TemplatePackage>(json);
                }

                if (package == null) return;

                var promptIdMap = new Dictionary<string, string>();
                var scriptIdMap = new Dictionary<string, string>();
                var templateIdMap = new Dictionary<string, string>();

                // 1. Process Prompts
                foreach (var prompt in package.Prompts)
                {
                    string oldId = prompt.PromptId;
                    string newId = oldId;
                    
                    var existing = _promptService.GetPrompts().FirstOrDefault(p => p.PromptId == prompt.PromptId || p.DisplayName == prompt.DisplayName);
                    if (existing != null)
                    {
                        var dialog = new ConflictResolutionDialog(prompt.DisplayName);
                        if (dialog.ShowDialog() == true)
                        {
                            switch (dialog.Result)
                            {
                                case ConflictResolutionDialog.Resolution.Overwrite:
                                    _promptService.AddOrUpdatePrompt(prompt);
                                    break;
                                case ConflictResolutionDialog.Resolution.Skip:
                                    newId = existing.PromptId;
                                    break;
                                case ConflictResolutionDialog.Resolution.Rename:
                                    prompt.DisplayName += "_imported";
                                    prompt.PromptId = Guid.NewGuid().ToString();
                                    newId = prompt.PromptId;
                                    _promptService.AddOrUpdatePrompt(prompt);
                                    break;
                            }
                        }
                    }
                    else
                    {
                        _promptService.AddOrUpdatePrompt(prompt);
                    }
                    promptIdMap[oldId] = newId;
                }

                // 2. Process Scripts
                foreach (var script in package.Scripts)
                {
                    string oldId = script.ScriptId;
                    string newId = oldId;

                    var existing = _scriptService.DiscoverScripts().FirstOrDefault(s => s.ScriptId == script.ScriptId || s.DisplayName == script.DisplayName);
                    string scriptFileName = Path.GetFileName(script.ScriptPath);
                    var scriptEntry = zip.GetEntry("scripts/" + scriptFileName);
                    string scriptContent = "";
                    if (scriptEntry != null)
                    {
                        using (var reader = new StreamReader(scriptEntry.Open()))
                        {
                            scriptContent = reader.ReadToEnd();
                        }
                    }

                    if (existing != null)
                    {
                        var dialog = new ConflictResolutionDialog(script.DisplayName);
                        if (dialog.ShowDialog() == true)
                        {
                            switch (dialog.Result)
                            {
                                case ConflictResolutionDialog.Resolution.Overwrite:
                                    _scriptService.SaveScript(scriptFileName, scriptContent);
                                    _scriptService.UpdateScriptMetadata(script);
                                    break;
                                case ConflictResolutionDialog.Resolution.Skip:
                                    newId = existing.ScriptId;
                                    break;
                                case ConflictResolutionDialog.Resolution.Rename:
                                    string ext = Path.GetExtension(scriptFileName);
                                    string nameNoExt = Path.GetFileNameWithoutExtension(scriptFileName);
                                    string newFileName = nameNoExt + "_imported" + ext;
                                    scriptFileName = newFileName;
                                    script.ScriptId = nameNoExt + "_imported";
                                    newId = script.ScriptId;
                                    
                                    _scriptService.SaveScript(scriptFileName, scriptContent);
                                    _scriptService.UpdateScriptMetadata(script);
                                    break;
                            }
                        }
                    }
                    else
                    {
                        _scriptService.SaveScript(scriptFileName, scriptContent);
                        _scriptService.UpdateScriptMetadata(script);
                    }
                    scriptIdMap[oldId] = newId;
                }

                // 3. Process Files
                string filesDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "OSEMAddIn", "TemplateFiles");
                Directory.CreateDirectory(filesDir);

                foreach (var kvp in package.Files)
                {
                    string fileName = kvp.Key;
                    string zipPath = kvp.Value;
                    var entry = zip.GetEntry(zipPath);
                    if (entry != null)
                    {
                        string destPath = Path.Combine(filesDir, fileName);
                        entry.ExtractToFile(destPath, true);
                    }
                }

                // 4. Process Templates
                foreach (var template in package.Templates)
                {
                    string oldId = template.TemplateId;
                    string newId = oldId;

                    // Fix Attachment Paths
                    var newPaths = new List<string>();
                    foreach(var p in template.AttachmentPaths)
                    {
                        string fName = Path.GetFileName(p);
                        if (package.Files.ContainsKey(fName))
                        {
                            newPaths.Add(Path.Combine(filesDir, fName));
                        }
                        else
                        {
                            newPaths.Add(p);
                        }
                    }
                    template.AttachmentPaths = newPaths;

                    var existing = _templateService.GetTemplates().FirstOrDefault(t => t.TemplateId == template.TemplateId || t.DisplayName == template.DisplayName);
                    if (existing != null)
                    {
                        var dialog = new ConflictResolutionDialog(template.DisplayName);
                        if (dialog.ShowDialog() == true)
                        {
                            switch (dialog.Result)
                            {
                                case ConflictResolutionDialog.Resolution.Overwrite:
                                    _templateService.AddOrUpdateTemplate(template);
                                    break;
                                case ConflictResolutionDialog.Resolution.Skip:
                                    newId = existing.TemplateId;
                                    break;
                                case ConflictResolutionDialog.Resolution.Rename:
                                    template.DisplayName += "_imported";
                                    template.TemplateId = Guid.NewGuid().ToString();
                                    newId = template.TemplateId;
                                    _templateService.AddOrUpdateTemplate(template);
                                    break;
                            }
                        }
                    }
                    else
                    {
                        _templateService.AddOrUpdateTemplate(template);
                    }
                    templateIdMap[oldId] = newId;
                }

                // 5. Post-Process: Fix Links
                foreach (var kvp in promptIdMap)
                {
                    string currentPromptId = kvp.Value;
                    var prompt = _promptService.GetPrompts().FirstOrDefault(p => p.PromptId == currentPromptId);
                    if (prompt != null && !string.IsNullOrEmpty(prompt.TemplateOverrideId))
                    {
                        var ids = prompt.TemplateOverrideId!.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(id => id.Trim())
                            .ToList();
                        
                        bool changed = false;
                        for (int i = 0; i < ids.Count; i++)
                        {
                            if (templateIdMap.ContainsKey(ids[i]))
                            {
                                ids[i] = templateIdMap[ids[i]];
                                changed = true;
                            }
                        }

                        if (changed)
                        {
                            prompt.TemplateOverrideId = string.Join(",", ids);
                            _promptService.AddOrUpdatePrompt(prompt);
                        }
                    }
                }

                foreach (var pkgScript in package.Scripts)
                {
                    if (!scriptIdMap.ContainsKey(pkgScript.ScriptId)) continue;

                    string finalScriptId = scriptIdMap[pkgScript.ScriptId];
                    var localScript = _scriptService.DiscoverScripts().FirstOrDefault(s => s.ScriptId == finalScriptId);

                    if (localScript != null)
                    {
                        var currentAssociations = new HashSet<string>(localScript.AssociatedTemplateIds);
                        bool changed = false;

                        foreach (var oldTid in pkgScript.AssociatedTemplateIds)
                        {
                            // If this template was part of the import package
                            if (templateIdMap.ContainsKey(oldTid))
                            {
                                string newTid = templateIdMap[oldTid];

                                // If the script has the old ID (because it was imported/overwritten), remove it
                                // Only remove if it's different (renamed)
                                if (oldTid != newTid && currentAssociations.Contains(oldTid))
                                {
                                    currentAssociations.Remove(oldTid);
                                    changed = true;
                                }

                                // Add the new ID
                                if (!currentAssociations.Contains(newTid))
                                {
                                    currentAssociations.Add(newTid);
                                    changed = true;
                                }
                            }
                        }

                        if (changed)
                        {
                            localScript.AssociatedTemplateIds = currentAssociations.ToList();
                            _scriptService.UpdateScriptMetadata(localScript);
                        }
                    }
                }
            }
        }
    }
}
