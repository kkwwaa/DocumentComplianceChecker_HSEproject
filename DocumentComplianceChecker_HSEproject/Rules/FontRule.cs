using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    public class FontRule : ValidationRule
    {
        public string RequiredFont { get; set; } = "Times New Roman";

        public override bool Validate(Paragraph paragraph, Run run = null)
        {
            var targetRun = run ?? paragraph.Elements<Run>().First();
            var font = targetRun.RunProperties?.RunFonts?.Ascii?.Value;
            return font == RequiredFont;
        }
    }
}