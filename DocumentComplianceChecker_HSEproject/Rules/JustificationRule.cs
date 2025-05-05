using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    public class JustificationRule : ValidationRule
    {
        public override bool Validate(Paragraph paragraph, Run run = null)
        {
            var justification = paragraph.ParagraphProperties?.Justification?.Val;

            // Проверка, задано ли выравнивание и является ли оно "both" (по ширине)
            return justification != null && justification.Value == JustificationValues.Both;
        }
    }
}