using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    public class ColorRule : ValidationRule
    {
        // Допустимые цвета — по умолчанию "auto" и "000000"
        public List<string> AllowedColors { get; set; } = new List<string> { "auto", "000000" };

        public override bool Validate(Paragraph paragraph, Run run = null)
        {
            var targetRun = run ?? paragraph.Elements<Run>().FirstOrDefault();
            if (targetRun == null) return true;

            var color = targetRun.RunProperties?.Color?.Val?.Value;

            // Если цвет не задан, считаем это "auto"
            var effectiveColor = string.IsNullOrEmpty(color) ? "auto" : color;

            return AllowedColors.Any(c =>
                effectiveColor.Equals(c, StringComparison.OrdinalIgnoreCase));
        }
    }
}
