using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Models
{
    // Класс шаблона, содержащий набор правил форматирования для проверки документа
    internal class Template { }
    //{
    //    // Правило проверки цвета шрифта (разрешён только белый)
    //    private class ColorRuleTemplate : IValidationRule
    //    {
    //        public string ErrorMessage { get; set; } = "Недопустимый цвет текста.";
    //        // Допустимые цвета (в данном случае только белый)
    //        public List<string> AllowedColors { get; set; } = new List<string> { "ffffff" };

    //        public bool RuleValidator(Paragraph paragraph, Run run = null)
    //        {
    //            // Получаем run (текстовый элемент), если не задан — берём первый в абзаце
    //            var targetRun = run ?? paragraph.Elements<Run>().FirstOrDefault();
    //            if (targetRun == null) return true; // Если run нет — считаем безошибочным

    //            // Получаем значение цвета шрифта
    //            var color = targetRun.RunProperties?.Color?.Val?.Value;

    //            // Если цвет не задан — считаем это ошибкой (можно изменить на допустимо)
    //            if (string.IsNullOrEmpty(color)) return false;

    //            // Сравниваем цвет с допустимыми (без учёта регистра)
    //            return AllowedColors.Any(c =>
    //                color.Equals(c, StringComparison.OrdinalIgnoreCase));
    //        }
    //    }

    //    // Правило выравнивания абзаца по центру
    //    private class JustificationRuleTemplate : IValidationRule
    //    {
    //        public string ErrorMessage { get; set; } = "Недопустимое выравнивание.";
    //        public bool RuleValidator(Paragraph paragraph, Run run = null)
    //        {
    //            // Получаем значение выравнивания
    //            var justification = paragraph.ParagraphProperties?.Justification?.Val;

    //            // Проверка: задано ли выравнивание и является ли оно по центру
    //            return justification != null && justification.Value == JustificationValues.Center;
    //        }
    //    }

    //    private class LineSpacingRuleTemplate : IValidationRule
    //    {
    //        public string ErrorMessage { get; set; } = "Недопустимый интервал.";
    //        internal double RequiredLineSpacing { get; set; } = 1.5;

    //        public bool RuleValidator(Paragraph paragraph, Run run = null)
    //        {
    //            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
    //            if (spacing == null)
    //                return false;

    //            // Значение в OpenXML задаётся в 1/20 пункта, 1.5 интервала ≈ 360
    //            if (int.TryParse(spacing.Line?.Value, out int lineValue))
    //            {
    //                double actualSpacing = lineValue / 20.0;
    //                return Math.Abs(actualSpacing - RequiredLineSpacing * 12) < 1; // 1.5 * 12 = 18pt
    //            }

    //            return false;
    //        }
    //    }

    //    private class FirstLineIndentRuleTemplate : IValidationRule
    //    {
    //        public string ErrorMessage { get; set; } = "Неверный отступ первой строки";
    //        public double RequiredIndentInCm { get; set; } = 1.25;

    //        public bool RuleValidator(Paragraph paragraph, Run run = null)
    //        {
    //            var firstLineIndent = paragraph.ParagraphProperties?.Indentation?.FirstLine;

    //            if (firstLineIndent == null)
    //                return false; // Если отступ не задан, считаем ошибкой

    //            // Преобразуем firstLineIndent в double, если это строка или другой тип
    //            double indentInPoints;
    //            if (!double.TryParse(firstLineIndent.ToString(), out indentInPoints))
    //                return false; // Если не удалось преобразовать, считаем ошибкой

    //            // Конвертируем отступ в см
    //            double indentInCm = indentInPoints / 567.0; // 1/567 части дюйма = 1 см

    //            return Math.Abs(indentInCm - RequiredIndentInCm) < 0.1; // Проверяем на допустимую погрешность
    //        }
    //    }

    //    private class PageMarginRuleTemplate : IValidationRule
    //    {
    //        public string ErrorMessage { get; set; } = "Недопустимые поля.";

    //        // Допустимые значения полей в сантиметрах
    //        private const double TopCm = 2.0;
    //        private const double BottomCm = 2.0;
    //        private const double LeftCm = 3.0;
    //        private const double RightCm = 1.5;

    //        // Преобразование см в twentieths of a point (единицы OpenXML)
    //        private static int CmToTwips(double cm) => (int)(cm * 567); // 1 см ≈ 567 twips

    //        public bool RuleValidator(Paragraph paragraph, Run run = null)
    //        {
    //            // Получаем доступ к документу через родительскую иерархию
    //            var document = paragraph.Ancestors<Body>()
    //                .FirstOrDefault()?.Parent as Document;
    //            if (document == null) return true;

    //            var sectionProps = paragraph.Descendants<SectionProperties>().FirstOrDefault()
    //                ?? document.Body?.Elements<SectionProperties>().FirstOrDefault();

    //            var pageMargin = sectionProps?.GetFirstChild<PageMargin>();
    //            if (pageMargin == null) return false;

    //            // Сравнение с допустимыми значениями
    //            return pageMargin.Top == CmToTwips(TopCm)
    //                && pageMargin.Bottom == CmToTwips(BottomCm)
    //                && pageMargin.Left == CmToTwips(LeftCm)
    //                && pageMargin.Right == CmToTwips(RightCm);
    //        }
    //    }

    //    private class ParagraphSpacingRuleTemplate : IValidationRule
    //    {
    //        public string ErrorMessage { get; set; } = "Недопустимые отступы.";
    //        public bool RuleValidator(Paragraph paragraph, Run run = null)
    //        {
    //            // Получаем свойства абзаца (ParagraphProperties)
    //            var props = paragraph.ParagraphProperties;
    //            if (props == null)
    //                return false; // Если свойства отсутствуют — считаем, что нарушено правило

    //            var indent = props.Indentation;              // Отступы абзаца (слева, справа и др.)
    //            var spacing = props.SpacingBetweenLines;     // Межабзацные интервалы (перед и после)

    //            // Проверяем, что отступ слева либо отсутствует, либо равен нулю
    //            bool leftOk = indent?.Left == null || indent.Left.Value == "0";

    //            // Проверяем, что отступ справа либо отсутствует, либо равен нулю
    //            bool rightOk = indent?.Right == null || indent.Right.Value == "0";

    //            // Проверяем, что интервал перед абзацем либо отсутствует, либо равен нулю
    //            bool beforeOk = spacing?.Before == null || spacing.Before.Value == "0";

    //            // Проверяем, что интервал после абзаца либо отсутствует, либо равен нулю
    //            bool afterOk = spacing?.After == null || spacing.After.Value == "0";

    //            // Абзац соответствует требованиям, только если все 4 условия выполнены
    //            return leftOk && rightOk && beforeOk && afterOk;
    //        }
    //    }

    //    private class ParagraphStyleAndSizeRuleTemplate : IValidationRule
    //    {
    //        public string ErrorMessage { get; set; } = "Неверный стиль или размер текста.";
    //        public bool RuleValidator(Paragraph paragraph, Run run = null)
    //        {
    //            // Получаем идентификатор стиля абзаца (например, "Заголовок 1", "Обычный")
    //            string styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

    //            // Если run не передан, берём первый Run в абзаце
    //            var targetRun = run ?? paragraph.Elements<Run>().FirstOrDefault();

    //            // Если run отсутствует или не содержит RunProperties — ошибка
    //            if (targetRun == null || targetRun.RunProperties == null)
    //                return false;

    //            // Получаем размер шрифта в полупунктах (например, "26" для 13 пт)
    //            var fontSizeStr = targetRun.RunProperties.FontSize?.Val?.Value;

    //            // Проверка наличия жирности (Bold)
    //            var isBold = targetRun.RunProperties.Bold != null;

    //            // Преобразуем строку размера в число
    //            if (!int.TryParse(fontSizeStr, out int sizeInHalfPoints))
    //                return false;

    //            // Перевод размера в пункты (делим на 2)
    //            double sizeInPoints = sizeInHalfPoints / 2.0;

    //            // Проверка на соответствие стилю, размеру и жирности
    //            switch (styleId)
    //            {
    //                // Заменить числа на числа из правил пользователя
    //                case "Заголовок 1":
    //                    return Math.Abs(sizeInPoints - 16) < 0.1 && isBold;
    //                case "Заголовок 2":
    //                    return Math.Abs(sizeInPoints - 14) < 0.1 && isBold;
    //                case "Заголовок 3":
    //                    return Math.Abs(sizeInPoints - 13) < 0.1 && isBold;
    //                case "Обычный":
    //                    return Math.Abs(sizeInPoints - 13) < 0.1;
    //                default:
    //                    return false; // Неизвестный или неподдерживаемый стиль
    //            }
    //        }
    //    }

    //        // Коллекция всех правил, применяемых в данном шаблоне
    //        public List<IValidationRule> Rules { get; } = new List<IValidationRule>();

    //    // Конструктор шаблона с добавлением всех правил в список
    //    internal Template()
    //    {
    //        // Добавление каждого из правил в шаблон
    //        Rules.Add(new ColorRuleTemplate());                        // Цвет шрифта: белый
    //        Rules.Add(new JustificationRuleTemplate());               // Абзац выровнен по центру
    //        Rules.Add(new LineSpacingRuleTemplate());                  // Межстрочный интервал: 1.5
    //        Rules.Add(new FirstLineIndentRuleTemplate());              // Отступ первой строки: 1.25 см
    //        Rules.Add(new PageMarginRuleTemplate());                   // Поля документа: верх/низ — 2 см, лево — 3 см, право — 1.5 см
    //        Rules.Add(new ParagraphSpacingRuleTemplate());             // Отступы до/после абзаца: 0
    //        Rules.Add(new ParagraphStyleAndSizeRuleTemplate());        // Стиль, размер и начертание абзаца
    //    }
    //}
}