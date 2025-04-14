using System;
using System.IO;
using System.Text.Json;
using DocumentComplianceChecker_HSEproject.Models;

namespace DocumentComplianceChecker_HSEproject.Services
{
    public class TemplateManager
    {
        public FormattingTemplate LoadTemplate(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("Template file not found", path);

            var json = File.ReadAllText(path);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };

            return JsonSerializer.Deserialize<FormattingTemplate>(json, options);
        }
    }
}
