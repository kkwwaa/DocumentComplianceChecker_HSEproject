using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Services;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    public class ColorRule : ValidationRule
    {
        private FormattingTemplate template;
        private Template template1;

        public ColorRule(FormattingTemplate template)
        {
            this.template = template;
        }

        public ColorRule(Template template1)
        {
            this.template1 = template1;
        }

        // Допустимые цвета: "auto" (по умолчанию) и черный ("000000")
        public List<string> AllowedColors { get; set; } = new() { "auto", "000000" };

        public override bool Validate(Paragraph paragraph, Run run = null)
        {
            // Получаем первый Run, если не передан напрямую
            var targetRun = run ?? paragraph.Descendants<Run>().FirstOrDefault();
            if (targetRun == null)
                return true; // Нет текста — нет ошибки

            // Пытаемся получить значение цвета
            var colorElement = targetRun.RunProperties?.Color;
            var colorValue = colorElement?.Val?.Value;

            // Если цвет не задан — считаем, что он "auto"
            var effectiveColor = string.IsNullOrWhiteSpace(colorValue) ? "auto" : colorValue;

            // Сравниваем с допустимыми цветами без учёта регистра
            return AllowedColors.Any(allowed =>
                string.Equals(effectiveColor, allowed, StringComparison.OrdinalIgnoreCase));
        }
    }
}
