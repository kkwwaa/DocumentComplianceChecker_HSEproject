using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Services;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    public class FontSizeRule : ValidationRule
    {
        private FormattingTemplate template;

        public FontSizeRule(FormattingTemplate template)
        {
            this.template = template;
        }

        public int MinSize { get; set; } = 12;
        public int MaxSize { get; set; } = 14;

        public override bool Validate(Paragraph paragraph, Run run = null)
        {
            var targetRun = run ?? paragraph.Elements<Run>().First();
            var fontSize = targetRun.RunProperties?.FontSize?.Val?.Value;

            if (fontSize == null)
                return false; // Если размер не указан - считаем ошибкой

            // Конвертируем строку в число (значение размера в OpenXML хранится в "полуточечных" единицах)
            if (int.TryParse(fontSize, out int sizeInHalfPoints))
            {
                // Конвертируем полуточечные единицы в обычные пункты (1 пункт = 2 полуточек)
                double sizeInPoints = sizeInHalfPoints / 2.0;
                return sizeInPoints >= MinSize && sizeInPoints <= MaxSize;
            }

            return false; // Если не удалось распарсить - считаем ошибкой
        }
    }
}
