using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Wordprocessing;

internal class BasicRules
{
    internal class ColorRule : IValidationRule
    {
        public string ErrorMessage { get; set; } = "Недопустимый цвет текста.";
        public List<string> AllowedColors { get; set; } = new() { "auto", "000000" };

        public bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            var targetRun = run ?? paragraph.Elements<Run>().FirstOrDefault();
            if (targetRun == null) return true;

            var color = targetRun.RunProperties?.Color?.Val?.Value;
            var effectiveColor = string.IsNullOrEmpty(color) ? "auto" : color;
            return AllowedColors.Any(c =>
                effectiveColor.Equals(c, StringComparison.OrdinalIgnoreCase));
        }
    }

    internal class FirstLineIndentRule : IValidationRule
    {
        public string ErrorMessage { get; set; } = "Неверный отступ первой строки.";
        public double RequiredIndentInCm { get; set; } = 1.25;

        public bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            var firstLineIndent = paragraph.ParagraphProperties?.Indentation?.FirstLine;
            if (firstLineIndent == null) return false;

            if (!double.TryParse(firstLineIndent.Value, out var indentTwips)) return false;
            var indentInCm = indentTwips / 567.0;
            return Math.Abs(indentInCm - RequiredIndentInCm) < 0.1;
        }
    }

    internal class JustificationRule : IValidationRule
    {
        public string ErrorMessage { get; set; } = "Недопустимое выравнивание.";
        public bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            var justification = paragraph.ParagraphProperties?.Justification?.Val;
            return justification != null && justification == JustificationValues.Both;
        }
    }

    internal class LineSpacingRule : IValidationRule
    {
        public string ErrorMessage { get; set; } = "Межстрочный интервал должен быть 1.5.";
        public bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            if (spacing?.Line == null) return false;

            if (int.TryParse(spacing.Line.Value, out int lineValue))
            {
                double actualSpacing = lineValue / 20.0;
                return Math.Abs(actualSpacing - 18.0) < 1.0; // 1.5 * 12 = 18 pt
            }
            return false;
        }
    }

    internal class PageMarginRule : IValidationRule
    {
        public string ErrorMessage { get; set; } = "Неверные поля страницы.";
        private static int CmToTwips(double cm) => (int)(cm * 567);

        public bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            var document = paragraph.Ancestors<Body>()?.FirstOrDefault()?.Parent as Document;
            if (document == null) return false;

            var sectionProps = paragraph.Descendants<SectionProperties>().FirstOrDefault()
                ?? document.Body?.Elements<SectionProperties>().FirstOrDefault();

            var margin = sectionProps?.GetFirstChild<PageMargin>();
            if (margin == null) return false;

            return margin.Top == CmToTwips(2.0)
                && margin.Bottom == CmToTwips(2.0)
                && margin.Left == CmToTwips(3.0)
                && margin.Right == CmToTwips(1.5)
                && margin.Header == CmToTwips(1.5)
                && margin.Footer == CmToTwips(1.25);
        }
    }

    internal class ParagraphSpacingRule : IValidationRule
    {
        public string ErrorMessage { get; set; } = "Недопустимые отступы абзаца.";

        public bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            var props = paragraph.ParagraphProperties;
            var indent = props?.Indentation;
            var spacing = props?.SpacingBetweenLines;

            bool leftOk = indent?.Left == null || indent.Left.Value == "0";
            bool rightOk = indent?.Right == null || indent.Right.Value == "0";
            bool beforeOk = spacing?.Before == null || spacing.Before.Value == "0";
            bool afterOk = spacing?.After == null || spacing.After.Value == "0";

            return leftOk && rightOk && beforeOk && afterOk;
        }
    }

    internal class ParagraphStyleAndSizeRule : IValidationRule
    {
        public string ErrorMessage { get; set; } = "Неверный стиль или размер текста.";

        public bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            string styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            var targetRun = run ?? paragraph.Elements<Run>().FirstOrDefault();
            if (targetRun?.RunProperties == null) return false;

            var fontName = targetRun.RunProperties.RunFonts?.Ascii?.Value;
            var sizeStr = targetRun.RunProperties.FontSize?.Val?.Value;
            var isBold = targetRun.RunProperties.Bold != null;

            if (!int.TryParse(sizeStr, out int sizeHalfPt)) return false;
            double sizePt = sizeHalfPt / 2.0;

            // Проверка обычного текста
            if (string.IsNullOrEmpty(styleId) || styleId == "Normal")
                return fontName == "Times New Roman" && Math.Abs(sizePt - 13) < 0.1 && !isBold;

            // Заголовки
            return styleId switch
            {
                "Heading1" => fontName == "Times New Roman" && isBold && Math.Abs(sizePt - 16) < 0.1,
                "Heading2" => fontName == "Times New Roman" && isBold && Math.Abs(sizePt - 14) < 0.1,
                "Heading3" => fontName == "Times New Roman" && isBold && Math.Abs(sizePt - 13) < 0.1,
                _ => false,
            };
        }
    }

    internal class HeadingStartsNewPageRule : IValidationRule
    {
        public string ErrorMessage { get; set; } = "Заголовок первого уровня должен начинаться с новой страницы.";

        public bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            string styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (styleId != "Heading1") return true;

            var breakBefore = paragraph.ParagraphProperties?.PageBreakBefore;
            return breakBefore != null;
        }
    }

    internal class HeadingSpacingRule : IValidationRule
    {
        public string ErrorMessage { get; set; } = "Неверные интервалы до/после заголовка.";

        public bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            string styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (styleId == null) return true;

            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            if (spacing == null) return false;

            var before = spacing.Before != null ? int.Parse(spacing.Before.Value) : 0;
            var after = spacing.After != null ? int.Parse(spacing.After.Value) : 0;

            return styleId switch
            {
                "Heading1" => before == 240 && after == 60, // 12pt, 3pt
                "Heading2" => before == 180 && after == 60, // 9pt, 3pt
                "Heading3" => before == 120 && after == 40, // 6pt, 2pt
                _ => true
            };
        }
    }

    internal class Heading3NotInTOCRule : IValidationRule
    {
        public string ErrorMessage { get; set; } = "Заголовки третьего уровня не должны отображаться в оглавлении.";

        public bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            // Определим, относится ли параграф к TOC
            string styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (styleId == null || !styleId.StartsWith("TOC")) return true;

            // Просто проверяем, если это заголовок 3 уровня
            return paragraph.InnerText.Contains("3 ");
        }
    }

}