using System;
Style&Review
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

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
            // 1. Проверка стиля параграфа
            var styleId = GetStyleId(paragraph);
            Console.WriteLine($"[Heading1Rule] Стиль параграфа: '{styleId}'");
            if (styleId != "2")
            {
                Console.WriteLine("[Heading1Rule] Стиль не Heading 1 (id ≠ 2). Пропускаем проверку параграфа.");
                return true;
            }

            // 2. ParagraphProperties должно присутствовать
            var props = paragraph.ParagraphProperties;
            if (props == null)
            {
                Console.WriteLine("[Heading1Rule] ParagraphProperties = null. Ошибка.");
                return false;
            }
            Console.WriteLine("[Heading1Rule] ParagraphProperties найдены.");

            // 3. Проверка spacing (пробелы до/после)
            var spacing = props.SpacingBetweenLines;
            if (spacing == null)
            {
                Console.WriteLine("[Heading1Rule] SpacingBetweenLines = null. Ошибка.");
                return false;
            }

            // raw-значения перед преобразованием
            string rawBefore = spacing.Before?.Value;
            string rawAfter = spacing.After?.Value;
            Console.WriteLine($"[Heading1Rule] Spacing.Before (raw): '{rawBefore}'");
            Console.WriteLine($"[Heading1Rule] Spacing.After  (raw): '{rawAfter}'");

            // парсим в int (твипсы)
            int before = int.TryParse(rawBefore, out var b) ? b : 0;
            int after = int.TryParse(rawAfter, out var a) ? a : 0;
            Console.WriteLine($"[Heading1Rule] Spacing.Before (parsed): {before} твипс (ожидается 0)");
            Console.WriteLine($"[Heading1Rule] Spacing.After  (parsed): {after} твипс (ожидается 240)");

            // Проверяем: до = 0, после = 240 (12 pt)
            if (before != 0 || after != 240)
            {
                Console.WriteLine("[Heading1Rule] Пробелы до/после не соответствуют требуемым.");
                return false;
            }

            // 4. Проверка отступа первой строки (Indentation.FirstLine)
            var indent = props.Indentation?.FirstLine;
            Console.WriteLine($"[Heading1Rule] Indentation.FirstLine: '{indent}' (ожидается '0' или отсутствует)");
            if (!string.IsNullOrEmpty(indent) && indent != "0")
            {
                Console.WriteLine("[Heading1Rule] Абзацный отступ первой строки не равен 0.");
                return false;
            }

            // 5. Проверка PageBreakBefore
            bool hasPageBreak = props.PageBreakBefore != null;
            Console.WriteLine($"[Heading1Rule] PageBreakBefore присутствует: {hasPageBreak} (ожидается true)");
            if (!hasPageBreak)
            {
                Console.WriteLine("[Heading1Rule] Заголовок не начинается с новой страницы.");
                return false;
            }

            Console.WriteLine("[Heading1Rule] Все проверки параграфа пройдены.");
            return true;
        }

        // Проверка свойств Run для стиля "Heading 1"
        public override bool ValidateRun(Paragraph paragraph, Run run)
        {
            // 1. Снова проверяем стиль параграфа
            var styleId = GetStyleId(paragraph);
            if (styleId != "2")
            {
                Console.WriteLine("[Heading1Rule] Стиль не Heading 1 (id ≠ 2). Пропускаем проверку Run.");
                return true;
            }

            // 2. Проверяем, что run не null
            if (run == null)
            {
                Console.WriteLine("[Heading1Rule] Run = null. Ошибка.");
                return false;
            }
            Console.WriteLine("[Heading1Rule] Run существует.");

            // 3. RunProperties должно присутствовать
            var runProps = run.RunProperties;
            if (runProps == null)
            {
                Console.WriteLine("[Heading1Rule] RunProperties = null. Ошибка.");
                return false;
            }
            Console.WriteLine("[Heading1Rule] RunProperties найдены.");

            // 4. Проверка шрифта
            var font = runProps.RunFonts?.Ascii?.Value;
            Console.WriteLine($"[Heading1Rule] RunFonts.Ascii: '{font}' (ожидается 'Times New Roman')");
            if (!string.Equals(font, "Times New Roman", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("[Heading1Rule] Шрифт не соответствует Times New Roman.");
                return false;
            }

            // 5. Проверка размера (FontSize.Val хранится в полутвипах)
            var sizeStr = runProps.FontSize?.Val?.Value;
            Console.WriteLine($"[Heading1Rule] FontSize.Val (raw): '{sizeStr}'");
            if (!int.TryParse(sizeStr, out var sizeHalfPt))
            {
                Console.WriteLine("[Heading1Rule] Не удалось распарсить FontSize.Val в int.");
                return false;
            }
            double sizePt = sizeHalfPt / 2.0;
            Console.WriteLine($"[Heading1Rule] Размер шрифта: {sizeHalfPt} (полутвипов) => {sizePt} pt (ожидается 16)");
            if (Math.Abs(sizePt - 16) > 0.1)
            {
                Console.WriteLine("[Heading1Rule] Размер шрифта отличается от 16 pt.");
                return false;
            }

            Console.WriteLine("[Heading1Rule] Все проверки Run пройдены.");
            return true;
        }
    }
}
