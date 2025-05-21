using DocumentFormat.OpenXml.Wordprocessing;
using DocumentComplianceChecker_HSEproject.Interfaces;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    // Класс для валидации стиля "Normal", реализует оба интерфейса
    internal class NormalStyleRule : StyleBasedRule
    {
        // Сообщение об ошибке, используется в ValidationResult
        public override string ErrorMessage => "Нарушения в обычном тексте.";

        // Проверка свойств параграфа для стиля "Normal"
        public override bool ValidateParagraph(Paragraph paragraph)
        {
            // Получаем стиль параграфа
            var styleId = GetStyleId(paragraph);
            Console.WriteLine($"[NormalStyleRule] Стиль параграфа: {styleId}");

            // Пропускаем, если стиль не "Normal"
            if (styleId != "Normal") return true;

            // Получаем свойства параграфа
            var props = paragraph.ParagraphProperties;
            if (props == null)
            {
                return false; // Свойства отсутствуют, возвращаем false
            }

            // Проверка межстрочного интервала (1.5 строки = 18 pt)
            var spacing = props.SpacingBetweenLines;
            if (spacing?.Line == null || !int.TryParse(spacing.Line.Value, out var lineVal)) return false;
            if (Math.Abs(lineVal / 20.0 - 18.0) > 1.0) return false;

            // Проверка красной строки (1.25 см)
            var indent = props.Indentation?.FirstLine;
            if (!double.TryParse(indent, out var indentTwips) || Math.Abs(indentTwips / 567.0 - 1.25) > 0.1) return false;

            // Проверка выравнивания (по ширине)
            var justification = props.Justification?.Val;
            if (justification != JustificationValues.Both) return false;

            // Проверка отступов (должны быть 0)
            bool leftOk = props.Indentation?.Left == null || props.Indentation.Left.Value == "0";
            bool rightOk = props.Indentation?.Right == null || props.Indentation.Right.Value == "0";
            bool beforeOk = spacing.Before == null || spacing.Before.Value == "0";
            bool afterOk = spacing.After == null || spacing.After.Value == "0";

            return leftOk && rightOk && beforeOk && afterOk;
        }

        // Проверка свойств Run для стиля "Normal"
        public override bool ValidateRun(Paragraph paragraph, Run run)
        {
            // Пропускаем, если стиль не "Normal"
            var styleId = GetStyleId(paragraph);
            if (styleId != "Normal") return true;

            // Проверяем наличие Run
            if (run == null) return false;

            // Получаем свойства Run
            var runProps = run.RunProperties;
            if (runProps == null) return false;

            // Проверка шрифта и размера (Times New Roman, 13 pt)
            var font = runProps.RunFonts?.Ascii?.Value;
            var sizeStr = runProps.FontSize?.Val?.Value;
            if (!int.TryParse(sizeStr, out var sizeHalfPt)) return false;
            double sizePt = sizeHalfPt / 2.0;

            if (font != "Times New Roman" || Math.Abs(sizePt - 13) > 0.1) return false;

            return true;
        }
    }
}