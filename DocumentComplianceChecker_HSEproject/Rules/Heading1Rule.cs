using DocumentFormat.OpenXml.Wordprocessing;
using DocumentComplianceChecker_HSEproject.Interfaces;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    // Класс для валидации стиля "Heading 1", реализует оба интерфейса
    internal class Heading1Rule : StyleBasedRule
    {
        // Сообщение об ошибке, используется в ValidationResult
        public override string ErrorMessage => "Нарушения в заголовке 1 уровня.";

        // Проверка свойств параграфа для стиля "Heading 1"
        public override bool ValidateParagraph(Paragraph paragraph)
        {
            // Получаем стиль параграфа
            var styleId = GetStyleId(paragraph);
            Console.WriteLine($"[Heading1Rule] Стиль параграфа: {styleId}");

            // Пропускаем, если стиль не "2" (Heading 1)
            if (styleId != "2") return true;

            // Получаем свойства параграфа
            var props = paragraph.ParagraphProperties;
            if (props == null) return false;

            // Проверка интервалов до и после (240 твипсов до, 60 твипсов после)
            var spacing = props.SpacingBetweenLines;
            if (spacing == null) return false;
            int before = int.TryParse(spacing.Before?.Value, out var b) ? b : 0;
            int after = int.TryParse(spacing.After?.Value, out var a) ? a : 0;
            if (before != 240 || after != 60) return false;

            // Проверка отсутствия абзацного отступа
            var indent = props.Indentation?.FirstLine;
            if (!string.IsNullOrEmpty(indent) && indent != "0") return false;

            // Проверка начала с новой страницы
            if (props.PageBreakBefore == null) return false;

            return true;
        }

        // Проверка свойств Run для стиля "Heading 1"
        public override bool ValidateRun(Paragraph paragraph, Run run)
        {
            // Пропускаем, если стиль не "2" (Heading 1)
            var styleId = GetStyleId(paragraph);
            if (styleId != "2") return true;

            // Проверяем наличие Run
            if (run == null) return false;

            // Получаем свойства Run
            var runProps = run.RunProperties;
            if (runProps == null) return false;

            // Проверка шрифта, размера и жирности (Times New Roman, 16 pt, жирный)
            var font = runProps.RunFonts?.Ascii?.Value;
            var sizeStr = runProps.FontSize?.Val?.Value;
            if (!int.TryParse(sizeStr, out var sizeHalfPt)) return false;
            double sizePt = sizeHalfPt / 2.0;

            if (font != "Times New Roman" || Math.Abs(sizePt - 16) > 0.1) return false;
            if (runProps.Bold == null) return false;

            return true;
        }
    }
}