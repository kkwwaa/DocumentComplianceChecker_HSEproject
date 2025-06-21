using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    // Класс для валидации стиля "Heading 3", реализует оба интерфейса
    internal class Heading3Rule : BasicRules
    {
        // Сообщение об ошибке, используется в ValidationResult
        public override string ErrorMessage => "Нарушения в заголовке 3 уровня.";

        // Проверка свойств параграфа для стиля "Heading 3"
        public override bool ValidateParagraph(Paragraph paragraph)
        {
            // Получаем стиль параграфа
            var styleId = GetStyleId(paragraph);
            Console.WriteLine($"[Heading3Rule] Стиль параграфа: {styleId}");

            // Пропускаем, если стиль не "4" (Heading 3)
            if (styleId != "4") return true;

            // Получаем свойства параграфа
            var props = paragraph.ParagraphProperties;
            if (props == null) return false;

            // Проверка интервалов до и после (120 твипсов до, 40 твипсов после)
            var spacing = props.SpacingBetweenLines;
            if (spacing == null) return false;
            int before = int.TryParse(spacing.Before?.Value, out var b) ? b : 0;
            int after = int.TryParse(spacing.After?.Value, out var a) ? a : 0;
            if (before != 120 || after != 40) return false;

            // Проверка отсутствия абзацного отступа
            var indent = props.Indentation?.FirstLine;
            if (!string.IsNullOrEmpty(indent) && indent != "0") return false;

            return true;
        }

        // Проверка свойств Run для стиля "Heading 3"
        public override bool ValidateRun(Paragraph paragraph, Run run)
        {
            // Пропускаем, если стиль не "4" (Heading 3)
            var styleId = GetStyleId(paragraph);
            if (styleId != "4") return true;

            // Проверяем наличие Run
            if (run == null) return false;

            // Получаем свойства Run
            var runProps = run.RunProperties;
            if (runProps == null) return false;

            // Проверка шрифта, размера и жирности (Times New Roman, 13 pt, жирный)
            var font = runProps.RunFonts?.Ascii?.Value;
            var sizeStr = runProps.FontSize?.Val?.Value;
            if (!int.TryParse(sizeStr, out var sizeHalfPt)) return false;
            double sizePt = sizeHalfPt / 2.0;

            if (font != "Times New Roman" || Math.Abs(sizePt - 13) > 0.1) return false;
            if (runProps.Bold == null) return false;

            return true;
        }
    }
}