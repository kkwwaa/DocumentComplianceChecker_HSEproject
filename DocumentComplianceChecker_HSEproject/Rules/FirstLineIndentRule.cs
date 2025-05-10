using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    internal class FirstLineIndentRule : ValidationRule
    {
        public double RequiredIndentInCm { get; set; } = 1.25;

        public override bool Validate(Paragraph paragraph, Run run = null)
        {
            var firstLineIndent = paragraph.ParagraphProperties?.Indentation?.FirstLine;

            if (firstLineIndent == null)
                return false; // Если отступ не задан, считаем ошибкой

            // Преобразуем firstLineIndent в double, если это строка или другой тип
            double indentInPoints;
            if (!double.TryParse(firstLineIndent.ToString(), out indentInPoints))
                return false; // Если не удалось преобразовать, считаем ошибкой

            // Конвертируем отступ в см
            double indentInCm = indentInPoints / 567.0; // 1/567 части дюйма = 1 см

            return Math.Abs(indentInCm - RequiredIndentInCm) < 0.1; // Проверяем на допустимую погрешность
        }
    }
}