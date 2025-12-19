using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal class BackupService
    {
        private readonly IEventRepository _eventRepository;
        private readonly DashboardTemplateService _templateService;
        private readonly PromptLibraryService _promptService;
        private readonly PythonScriptService _scriptService;
        private readonly EmailTemplateService _emailTemplateService;

        public BackupService(
            IEventRepository eventRepository,
            DashboardTemplateService templateService,
            PromptLibraryService promptService,
            PythonScriptService scriptService,
            EmailTemplateService emailTemplateService)
        {
            _eventRepository = eventRepository;
            _templateService = templateService;
            _promptService = promptService;
            _scriptService = scriptService;
            _emailTemplateService = emailTemplateService;
        }

        public async Task ExportBackupAsync(string filePath)
        {
            var package = new BackupPackage();
            
            // 1. Collect Data
            package.Events = (await _eventRepository.GetAllAsync()).ToList();
            package.DashboardTemplates = _templateService.GetTemplates().ToList();
            package.Prompts = _promptService.GetPrompts().ToList();
            package.Scripts = _scriptService.DiscoverScripts().ToList();
            
            // Collect Email Templates (Compose and Reply)
            var composeTemplates = _emailTemplateService.GetTemplates(EmailTemplateType.Compose);
            var replyTemplates = _emailTemplateService.GetTemplates(EmailTemplateType.Reply);
            package.EmailTemplates.AddRange(composeTemplates);
            package.EmailTemplates.AddRange(replyTemplates);

            // 2. Collect Files
            // Template Attachments
            foreach (var template in package.DashboardTemplates)
            {
                foreach (var attachmentPath in template.AttachmentPaths)
                {
                    if (File.Exists(attachmentPath))
                    {
                        string fileName = Path.GetFileName(attachmentPath);
                        string key = "attachments/" + fileName;
                        if (!package.Files.ContainsKey(key))
                        {
                            package.Files[key] = key;
                        }
                    }
                }
            }

            // Scripts
            foreach (var script in package.Scripts)
            {
                if (File.Exists(script.ScriptPath))
                {
                    string fileName = Path.GetFileName(script.ScriptPath);
                    string key = "scripts/" + fileName;
                    if (!package.Files.ContainsKey(key))
                    {
                        package.Files[key] = key;
                    }
                }
            }

            // 3. Create Zip
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var zip = ZipFile.Open(filePath, ZipArchiveMode.Create))
            {
                // Write Manifest
                string json = JsonConvert.SerializeObject(package, Formatting.Indented);
                var entry = zip.CreateEntry("backup_manifest.json");
                using (var writer = new StreamWriter(entry.Open()))
                {
                    await writer.WriteAsync(json);
                }

                // Write Files
                foreach (var kvp in package.Files)
                {
                    string key = kvp.Key; // e.g. "scripts/foo.py" or "attachments/bar.pdf"
                    string zipPath = kvp.Value; // same as key

                    string? sourcePath = null;

                    if (key.StartsWith("scripts/"))
                    {
                        // It's a script
                        string scriptName = Path.GetFileName(key);
                        var script = package.Scripts.FirstOrDefault(s => Path.GetFileName(s.ScriptPath) == scriptName);
                        if (script != null) sourcePath = script.ScriptPath;
                    }
                    else if (key.StartsWith("attachments/"))
                    {
                        // It's an attachment
                        string fileName = Path.GetFileName(key);
                        foreach(var t in package.DashboardTemplates)
                        {
                            var match = t.AttachmentPaths.FirstOrDefault(p => Path.GetFileName(p) == fileName);
                            if (match != null) { sourcePath = match; break; }
                        }
                    }

                    if (sourcePath != null && File.Exists(sourcePath))
                    {
                        zip.CreateEntryFromFile(sourcePath, zipPath);
                    }
                }
            }
        }

        public async Task ImportBackupAsync(string filePath)
        {
            using (var zip = ZipFile.OpenRead(filePath))
            {
                var manifestEntry = zip.GetEntry("backup_manifest.json");
                if (manifestEntry == null) throw new Exception("Invalid backup file: missing backup_manifest.json");

                BackupPackage? package;
                using (var reader = new StreamReader(manifestEntry.Open()))
                {
                    string json = await reader.ReadToEndAsync();
                    package = JsonConvert.DeserializeObject<BackupPackage>(json);
                }

                if (package == null) return;

                // 1. Restore Files
                string defaultScriptFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "OSEM", "Scripts");
                Directory.CreateDirectory(defaultScriptFolder);

                string defaultAttachmentFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "OSEM", "Attachments");
                Directory.CreateDirectory(defaultAttachmentFolder);

                foreach (var kvp in package.Files)
                {
                    string key = kvp.Key;
                    string zipPath = kvp.Value;
                    var entry = zip.GetEntry(zipPath);
                    if (entry != null)
                    {
                        string targetPath;
                        if (key.StartsWith("scripts/"))
                        {
                            targetPath = Path.Combine(defaultScriptFolder, Path.GetFileName(key));
                        }
                        else
                        {
                            targetPath = Path.Combine(defaultAttachmentFolder, Path.GetFileName(key));
                        }

                        // Overwrite if exists
                        entry.ExtractToFile(targetPath, true);
                    }
                }

                // 2. Merge Data
                
                // Events
                foreach (var evt in package.Events)
                {
                    // Import (Add or Update)
                    await _eventRepository.ImportAsync(evt);
                }
                
                // Templates
                foreach (var tmpl in package.DashboardTemplates)
                {
                    // Fix attachment paths
                    var newPaths = new List<string>();
                    foreach(var p in tmpl.AttachmentPaths)
                    {
                        string fileName = Path.GetFileName(p);
                        string newPath = Path.Combine(defaultAttachmentFolder, fileName);
                        newPaths.Add(newPath);
                    }
                    tmpl.AttachmentPaths = newPaths;

                    _templateService.AddOrUpdateTemplate(tmpl);
                }

                // Prompts
                foreach (var prompt in package.Prompts)
                {
                    _promptService.AddOrUpdatePrompt(prompt);
                }

                // Scripts
                foreach (var script in package.Scripts)
                {
                    string fileName = Path.GetFileName(script.ScriptPath);
                    script.ScriptPath = Path.Combine(defaultScriptFolder, fileName);
                    
                    _scriptService.UpdateScriptMetadata(script);
                }
                
                // Email Templates
                foreach (var emailTmpl in package.EmailTemplates)
                {
                    _emailTemplateService.SaveTemplate(emailTmpl);
                }
            }
        }
    }
}
