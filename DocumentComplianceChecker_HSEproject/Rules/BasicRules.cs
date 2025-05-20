using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentFormat.OpenXml.Wordprocessing;

internal class BasicRules
{
    internal abstract class StyleBasedRule : IValidationRule
    {
        public abstract string ErrorMessage { get; }
        public abstract bool RuleValidator(Paragraph paragraph, Run run = null);

        protected string GetStyleId(Paragraph paragraph)
        {
            return paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal";
        }

        protected Run GetTargetRun(Paragraph paragraph, Run run)
        {
            return run ?? paragraph.Elements<Run>().FirstOrDefault();
        }
    }

    internal class NormalStyleRule : StyleBasedRule
    {
        public override string ErrorMessage => "Нарушения в обычном тексте.";

        public override bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            var styleId = GetStyleId(paragraph);
            Console.WriteLine($"[NormalStyleRule] Стиль параграфа: {styleId}");

            if (styleId != "Normal") return true;

            var props = paragraph.ParagraphProperties;
            var runProps = GetTargetRun(paragraph, run)?.RunProperties;
            if (props == null || runProps == null) return false;

            // Шрифт
            var font = runProps.RunFonts?.Ascii?.Value;
            var sizeStr = runProps.FontSize?.Val?.Value;
            if (!int.TryParse(sizeStr, out var sizeHalfPt)) return false;
            double sizePt = sizeHalfPt / 2.0;
            if (font != "Times New Roman" || Math.Abs(sizePt - 13) > 0.1) return false;

            // Межстрочный интервал
            var spacing = props.SpacingBetweenLines;
            if (spacing?.Line == null || !int.TryParse(spacing.Line.Value, out var lineVal)) return false;
            if (Math.Abs(lineVal / 20.0 - 18.0) > 1.0) return false;

            // Красная строка
            var indent = props.Indentation?.FirstLine;
            if (!double.TryParse(indent, out var indentTwips) || Math.Abs(indentTwips / 567.0 - 1.25) > 0.1) return false;

            // Выравнивание
            var justification = props.Justification?.Val;
            if (justification != JustificationValues.Both) return false;

            // Параметры абзацев
            bool leftOk = props.Indentation?.Left == null || props.Indentation.Left.Value == "0";
            bool rightOk = props.Indentation?.Right == null || props.Indentation.Right.Value == "0";
            bool beforeOk = spacing.Before == null || spacing.Before.Value == "0";
            bool afterOk = spacing.After == null || spacing.After.Value == "0";

            return leftOk && rightOk && beforeOk && afterOk;
        }
    }

    internal class Heading1Rule : StyleBasedRule
    {
        public override string ErrorMessage => "Нарушения в заголовке 1 уровня.";

        public override bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            var styleId = GetStyleId(paragraph);
            Console.WriteLine($"[Heading1Rule] Стиль параграфа: {styleId}");

            if (styleId != "2") return true;
            var props = paragraph.ParagraphProperties;
            var runProps = GetTargetRun(paragraph, run)?.RunProperties;
            if (props == null || runProps == null) return false;

            // Шрифт и размер
            var font = runProps.RunFonts?.Ascii?.Value;
            var sizeStr = runProps.FontSize?.Val?.Value;
            if (!int.TryParse(sizeStr, out var sizeHalfPt)) return false;
            double sizePt = sizeHalfPt / 2.0;
            if (font != "Times New Roman" || Math.Abs(sizePt - 16) > 0.1 || runProps.Bold == null) return false;

            // Интервалы до и после
            var spacing = props.SpacingBetweenLines;
            if (spacing == null) return false;
            int before = int.TryParse(spacing.Before?.Value, out var b) ? b : 0;
            int after = int.TryParse(spacing.After?.Value, out var a) ? a : 0;
            if (before != 240 || after != 60) return false;

            // Нет абзацного отступа
            var indent = props.Indentation?.FirstLine;
            if (!string.IsNullOrEmpty(indent) && indent != "0") return false;

            // Начало с новой страницы
            return props.PageBreakBefore != null;
        }
    }

    internal class Heading2Rule : StyleBasedRule
    {
        public override string ErrorMessage => "Нарушения в заголовке 2 уровня.";

        public override bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            var styleId = GetStyleId(paragraph);
            Console.WriteLine($"[Heading2Rule] Стиль параграфа: {styleId}");

            if (styleId != "3") return true;
            var props = paragraph.ParagraphProperties;
            var runProps = GetTargetRun(paragraph, run)?.RunProperties;
            if (props == null || runProps == null) return false;

            var font = runProps.RunFonts?.Ascii?.Value;
            var sizeStr = runProps.FontSize?.Val?.Value;
            if (!int.TryParse(sizeStr, out var sizeHalfPt)) return false;
            double sizePt = sizeHalfPt / 2.0;
            if (font != "Times New Roman" || Math.Abs(sizePt - 14) > 0.1 || runProps.Bold == null) return false;

            var spacing = props.SpacingBetweenLines;
            if (spacing == null) return false;
            int before = int.TryParse(spacing.Before?.Value, out var b) ? b : 0;
            int after = int.TryParse(spacing.After?.Value, out var a) ? a : 0;
            if (before != 180 || after != 60) return false;

            var indent = props.Indentation?.FirstLine;
            return string.IsNullOrEmpty(indent) || indent == "0";
        }
    }

    internal class Heading3Rule : StyleBasedRule
    {
        public override string ErrorMessage => "Нарушения в заголовке 3 уровня.";

        public override bool RuleValidator(Paragraph paragraph, Run run = null)
        {
            var styleId = GetStyleId(paragraph);
            Console.WriteLine($"[Heading3Rule] Стиль параграфа: {styleId}");

            if (styleId != "4") return true;
            var props = paragraph.ParagraphProperties;
            var runProps = GetTargetRun(paragraph, run)?.RunProperties;
            if (props == null || runProps == null) return false;

            var font = runProps.RunFonts?.Ascii?.Value;
            var sizeStr = runProps.FontSize?.Val?.Value;
            if (!int.TryParse(sizeStr, out var sizeHalfPt)) return false;
            double sizePt = sizeHalfPt / 2.0;
            if (font != "Times New Roman" || Math.Abs(sizePt - 13) > 0.1 || runProps.Bold == null) return false;

            var spacing = props.SpacingBetweenLines;
            if (spacing == null) return false;
            int before = int.TryParse(spacing.Before?.Value, out var b) ? b : 0;
            int after = int.TryParse(spacing.After?.Value, out var a) ? a : 0;
            if (before != 120 || after != 40) return false;

            var indent = props.Indentation?.FirstLine;
            if (!string.IsNullOrEmpty(indent) && indent != "0") return false;

            // Не входит в оглавление
            var isInTOC = paragraph.Ancestors<SdtBlock>()
                .Any(sdt => sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == "TOC");

            return !isInTOC;
        }
    }

    internal class PageMarginRule : IValidationRule
    {
        public string ErrorMessage => "Неверные поля страницы.";
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
}
