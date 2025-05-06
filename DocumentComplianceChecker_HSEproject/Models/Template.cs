using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Rules;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Models
{
    // Класс шаблона, содержащий набор правил форматирования для проверки документа
    internal class Template
    {
        // Правило проверки цвета шрифта (разрешён только белый)
        public class ColorRule1 : ValidationRule
        {
            // Допустимые цвета (в данном случае только белый)
            public List<string> AllowedColors { get; set; } = new List<string> { "ffffff" };

            public override bool Validate(Paragraph paragraph, Run run = null)
            {
                // Получаем run (текстовый элемент), если не задан — берём первый в абзаце
                var targetRun = run ?? paragraph.Elements<Run>().FirstOrDefault();
                if (targetRun == null) return true; // Если run нет — считаем безошибочным

                // Получаем значение цвета шрифта
                var color = targetRun.RunProperties?.Color?.Val?.Value;

                // Если цвет не задан — считаем это ошибкой (можно изменить на допустимо)
                if (string.IsNullOrEmpty(color)) return false;

                // Сравниваем цвет с допустимыми (без учёта регистра)
                return AllowedColors.Any(c =>
                    color.Equals(c, StringComparison.OrdinalIgnoreCase));
            }
        }

        // Правило выравнивания абзаца по центру
        public class JustificationRule1 : ValidationRule
        {
            public override bool Validate(Paragraph paragraph, Run run = null)
            {
                // Получаем значение выравнивания
                var justification = paragraph.ParagraphProperties?.Justification?.Val;

                // Проверка: задано ли выравнивание и является ли оно по центру
                return justification != null && justification.Value == JustificationValues.Center;
            }
        }

        // Коллекция всех правил, применяемых в данном шаблоне
        public List<ValidationRule> Rules { get; } = new List<ValidationRule>();

        // Конструктор шаблона с добавлением всех правил в список
        public Template()
        {
            // Добавление каждого из правил в шаблон
            Rules.Add(new ColorRule1());                        // Цвет шрифта: белый
            Rules.Add(new JustificationRule1());               // Абзац выровнен по центру
            Rules.Add(new LineSpacingRule());                  // Межстрочный интервал: 1.5
            Rules.Add(new FirstLineIndentRule());              // Отступ первой строки: 1.25 см
            Rules.Add(new PageMarginRule());                   // Поля документа: верх/низ — 2 см, лево — 3 см, право — 1.5 см
            Rules.Add(new ParagraphSpacingRule());             // Отступы до/после абзаца: 0
            Rules.Add(new ParagraphStyleAndSizeRule());        // Стиль, размер и начертание абзаца
        }
    }
}
