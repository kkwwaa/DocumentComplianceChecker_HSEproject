using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    // Правило для проверки соответствия стиля абзаца, размера шрифта и жирности текста
    public class ParagraphStyleAndSizeRule : ValidationRule
    {
        public override bool Validate(Paragraph paragraph, Run run = null)
        {
            // Получаем идентификатор стиля абзаца (например, "Заголовок 1", "Обычный")
            string styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

            // Если run не передан, берём первый Run в абзаце
            var targetRun = run ?? paragraph.Elements<Run>().FirstOrDefault();

            // Если run отсутствует или не содержит RunProperties — ошибка
            if (targetRun == null || targetRun.RunProperties == null)
                return false;

            // Получаем размер шрифта в полупунктах (например, "26" для 13 пт)
            var fontSizeStr = targetRun.RunProperties.FontSize?.Val?.Value;

            // Проверка наличия жирности (Bold)
            var isBold = targetRun.RunProperties.Bold != null;

            // Преобразуем строку размера в число
            if (!int.TryParse(fontSizeStr, out int sizeInHalfPoints))
                return false;

            // Перевод размера в пункты (делим на 2)
            double sizeInPoints = sizeInHalfPoints / 2.0;

            // Проверка на соответствие стилю, размеру и жирности
            switch (styleId)
            {
                case "Заголовок 1":
                    return Math.Abs(sizeInPoints - 16) < 0.1 && isBold;
                case "Заголовок 2":
                    return Math.Abs(sizeInPoints - 14) < 0.1 && isBold;
                case "Заголовок 3":
                    return Math.Abs(sizeInPoints - 13) < 0.1 && isBold;
                case "Обычный":
                    return Math.Abs(sizeInPoints - 13) < 0.1;
                default:
                    return false; // Неизвестный или неподдерживаемый стиль
            }
        }
    }
}
