using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal sealed class EmailTemplateService
    {
        private readonly string _storePath;
        private readonly List<EmailTemplate> _templates = new();
        private readonly JsonSerializerSettings _serializerSettings = new()
        {
            Formatting = Formatting.Indented,
            TypeNameHandling = TypeNameHandling.None
        };

        public EmailTemplateService(string? storePath = null)
        {
            _storePath = storePath ?? BuildDefaultStorePath();
            Directory.CreateDirectory(Path.GetDirectoryName(_storePath)!);
            LoadFromDisk();
            EnsureSeedTemplates();
        }

        public IReadOnlyList<EmailTemplate> GetTemplates(EmailTemplateType type)
        {
            return _templates
                .Where(t => t.TemplateType == type)
                .Select(Clone)
                .ToList();
        }

        public void SaveTemplate(EmailTemplate template)
        {
            if (template is null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            var existing = _templates.FirstOrDefault(t => string.Equals(t.TemplateId, template.TemplateId, StringComparison.OrdinalIgnoreCase));
            if (existing is null)
            {
                _templates.Add(Clone(template));
            }
            else
            {
                existing.DisplayName = template.DisplayName;
                existing.Subject = template.Subject;
                existing.Body = template.Body;
                existing.TemplateType = template.TemplateType;
            }

            Persist();
        }

        public void DeleteTemplate(string templateId)
        {
            if (string.IsNullOrWhiteSpace(templateId))
            {
                return;
            }

            var removed = _templates.RemoveAll(t => string.Equals(t.TemplateId, templateId, StringComparison.OrdinalIgnoreCase));
            if (removed > 0)
            {
                Persist();
            }
        }

        private void LoadFromDisk()
        {
            if (!File.Exists(_storePath))
            {
                return;
            }

            var json = File.ReadAllText(_storePath);
            var items = JsonConvert.DeserializeObject<List<EmailTemplate>>(json, _serializerSettings);
            if (items is null)
            {
                return;
            }

            _templates.Clear();
            _templates.AddRange(items);
        }

        private void Persist()
        {
            var json = JsonConvert.SerializeObject(_templates, _serializerSettings);
            File.WriteAllText(_storePath, json);
        }

        private void EnsureSeedTemplates()
        {
            if (_templates.Count > 0)
            {
                return;
            }

            var replyTemplate = new EmailTemplate
            {
                DisplayName = Properties.Resources.Standard_Reply_Template,
                Subject = string.Empty,
                Body = Properties.Resources.p_Hello_p_p_We_have_received_y_5e6bbc,
                TemplateType = EmailTemplateType.Reply
            };

            var composeTemplate = new EmailTemplate
            {
                DisplayName = Properties.Resources.New_Event_Notification,
                Subject = Properties.Resources.Latest_update_on_Event_EventId,
                Body = Properties.Resources.p_Dear_Customer_p_p_Here_is_th_713eca,
                TemplateType = EmailTemplateType.Compose
            };

            _templates.Add(replyTemplate);
            _templates.Add(composeTemplate);
            Persist();
        }

        private static EmailTemplate Clone(EmailTemplate template)
        {
            return new EmailTemplate
            {
                TemplateId = template.TemplateId,
                DisplayName = template.DisplayName,
                Subject = template.Subject,
                Body = template.Body,
                TemplateType = template.TemplateType
            };
        }

        private static string BuildDefaultStorePath()
        {
            var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(root, "OSEM", "email-templates.json");
        }
    }
}
