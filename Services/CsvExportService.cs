using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal sealed class CsvExportService
    {
    public Task<string> ExportAsync(IEnumerable<EventRecord> eventsToExport, DashboardTemplate template, string targetDirectory)
        {
            if (eventsToExport is null)
            {
                throw new ArgumentNullException(nameof(eventsToExport));
            }

            if (template is null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            Directory.CreateDirectory(targetDirectory);
            var fileName = $"OSEM-{template.TemplateId}-{DateTime.UtcNow:yyyyMMdd-HHmmss}.csv";
            var path = Path.Combine(targetDirectory, fileName);

            var builder = new StringBuilder();
            builder.AppendLine(string.Join(",", template.Fields));

            foreach (var record in eventsToExport)
            {
                var line = template.Fields.Select(field => Escape(GetDashboardValue(record, field))).ToArray();
                builder.AppendLine(string.Join(",", line));
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
            return Task.FromResult(path);
        }

        public Task<string> ExportAsync(IEnumerable<EventRecord> eventsToExport, DashboardTemplate template, string targetDirectory, string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return ExportAsync(eventsToExport, template, targetDirectory);
            }

            Directory.CreateDirectory(targetDirectory);
            var path = Path.Combine(targetDirectory, fileName);
            var builder = new StringBuilder();
            builder.AppendLine(string.Join(",", template.Fields));

            foreach (var record in eventsToExport)
            {
                var line = template.Fields.Select(field => Escape(GetDashboardValue(record, field))).ToArray();
                builder.AppendLine(string.Join(",", line));
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
            return Task.FromResult(path);
        }

        private static string GetDashboardValue(EventRecord record, string field)
        {
            var item = record.DashboardItems.FirstOrDefault(i => string.Equals(i.Key, field, StringComparison.OrdinalIgnoreCase));
            return item?.Value ?? string.Empty;
        }

        private static string Escape(string value)
        {
            if (value.Contains('"') || value.Contains(',') || value.Contains('\n'))
            {
                return $"\"{value.Replace("\"", "\"\"")}\"";
            }

            return value;
        }
    }
}
