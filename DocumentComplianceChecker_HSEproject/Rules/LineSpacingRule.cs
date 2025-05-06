using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    public class LineSpacingRule : ValidationRule
    {
        public double RequiredLineSpacing { get; set; } = 1.5;

        public override bool Validate(Paragraph paragraph, Run run = null)
        {
            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            if (spacing == null)
                return false;

            // Значение в OpenXML задаётся в 1/20 пункта, 1.5 интервала ≈ 360
            if (int.TryParse(spacing.Line?.Value, out int lineValue))
            {
                double actualSpacing = lineValue / 20.0;
                return Math.Abs(actualSpacing - RequiredLineSpacing * 12) < 1; // 1.5 * 12 = 18pt
            }

            return false;
        }
    }
}