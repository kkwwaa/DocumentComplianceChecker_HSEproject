using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    // Класс для валидации стиля "Heading 2", реализует оба интерфейса
    internal class Heading2Rule : BasicRules
    {
        public override string ErrorMessage => "Нарушения в заголовке 2 уровня.";

        // Проверка свойств параграфа для стиля "Heading 2"
        public override bool ValidateParagraph(Paragraph paragraph)
        {
            var styleId = GetStyleId(paragraph);
            Console.WriteLine($"[Heading2Rule] Стиль параграфа: {styleId}");

            if (styleId != "3")
            {
                Console.WriteLine("[Heading2Rule] Стиль не Heading 2 (id ≠ 3). Пропускаем проверку параграфа.");
                return true;
            }

            var props = paragraph.ParagraphProperties;
            if (props == null)
            {
                Console.WriteLine("[Heading2Rule] ОШИБКА: ParagraphProperties = null.");
                return false;
            }
            Console.WriteLine("[Heading2Rule] ParagraphProperties найдены.");

            var spacing = props.SpacingBetweenLines;
            if (spacing == null)
            {
                Console.WriteLine("[Heading2Rule] ОШИБКА: SpacingBetweenLines = null.");
                return false;
            }
            Console.WriteLine("[Heading2Rule] SpacingBetweenLines найден.");

            string rawBefore = spacing.Before?.Value;
            string rawAfter = spacing.After?.Value;
            Console.WriteLine($"[Heading2Rule] Spacing.Before (raw): '{rawBefore}'");
            Console.WriteLine($"[Heading2Rule] Spacing.After  (raw): '{rawAfter}'");

            int before = int.TryParse(rawBefore, out var b) ? b : int.MinValue;
            int after = int.TryParse(rawAfter, out var a) ? a : int.MinValue;
            Console.WriteLine($"[Heading2Rule] Spacing.Before (parsed): {before}");
            Console.WriteLine($"[Heading2Rule] Spacing.After  (parsed): {after}");

            if (before != 180 || after != 60)
            {
                Console.WriteLine($"[Heading2Rule] ОШИБКА: ожидалось Before=180, After=60; получили Before={before}, After={after}.");
                return false;
            }
            Console.WriteLine("[Heading2Rule] Интервалы до/после соответствуют требованиям.");

            var indent = props.Indentation?.FirstLine;
            Console.WriteLine($"[Heading2Rule] Indentation.FirstLine: '{indent}'");
            if (!string.IsNullOrEmpty(indent) && indent != "0")
            {
                Console.WriteLine($"[Heading2Rule] ОШИБКА: Indentation.FirstLine должно отсутствовать или быть '0', а найдено '{indent}'.");
                return false;
            }
            Console.WriteLine("[Heading2Rule] Абзацный отступ первой строки соответствует требованиям.");

            Console.WriteLine("[Heading2Rule] Все проверки параграфа пройдены успешно.");
            return true;
        }

        // Проверка свойств Run для стиля "Heading 2"
        public override bool ValidateRun(Paragraph paragraph, Run run)
        {
            var styleId = GetStyleId(paragraph);
            Console.WriteLine($"[Heading2Rule] (Run) Стиль параграфа: {styleId}");

            if (styleId != "3")
            {
                Console.WriteLine("[Heading2Rule] (Run) Стиль не Heading 2 (id ≠ 3). Пропускаем проверку Run.");
                return true;
            }

            if (run == null)
            {
                Console.WriteLine("[Heading2Rule] ОШИБКА: Run = null.");
                return false;
            }
            Console.WriteLine("[Heading2Rule] Run существует.");

            var runProps = run.RunProperties;
            if (runProps == null)
            {
                Console.WriteLine("[Heading2Rule] ОШИБКА: RunProperties = null.");
                return false;
            }
            Console.WriteLine("[Heading2Rule] RunProperties найдены.");

            var font = runProps.RunFonts?.Ascii?.Value;
            Console.WriteLine($"[Heading2Rule] RunFonts.Ascii: '{font}' (ожидается 'Times New Roman')");
            if (!string.Equals(font, "Times New Roman", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"[Heading2Rule] ОШИБКА: шрифт должен быть 'Times New Roman', а найден '{font}'.");
                return false;
            }
            Console.WriteLine("[Heading2Rule] Шрифт соответствует требованиям.");

            var sizeStr = runProps.FontSize?.Val?.Value;
            Console.WriteLine($"[Heading2Rule] FontSize.Val (raw): '{sizeStr}'");
            if (!int.TryParse(sizeStr, out var sizeHalfPt))
            {
                Console.WriteLine($"[Heading2Rule] ОШИБКА: не удалось распознать FontSize.Val='{sizeStr}'.");
                return false;
            }
            double sizePt = sizeHalfPt / 2.0;
            Console.WriteLine($"[Heading2Rule] Размер шрифта: {sizePt} pt (полутвипов {sizeHalfPt}) (ожидается 14 pt)");
            if (Math.Abs(sizePt - 14) > 0.1)
            {
                Console.WriteLine($"[Heading2Rule] ОШИБКА: размер шрифта должен быть 14 pt, а найден {sizePt} pt.");
                return false;
            }
            Console.WriteLine("[Heading2Rule] Размер шрифта соответствует требованиям.");

            Console.WriteLine("[Heading2Rule] Все проверки Run пройдены успешно.");
            return true;
        }
    }
}
