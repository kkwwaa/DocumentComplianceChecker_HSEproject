using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    // Правило для проверки отсутствия отступов и интервалов у абзацев
    internal class ParagraphSpacingRule : ValidationRule
    {
        public override bool Validate(Paragraph paragraph, Run run = null)
        {
            // Получаем свойства абзаца (ParagraphProperties)
            var props = paragraph.ParagraphProperties;
            if (props == null)
                return false; // Если свойства отсутствуют — считаем, что нарушено правило

            var indent = props.Indentation;              // Отступы абзаца (слева, справа и др.)
            var spacing = props.SpacingBetweenLines;     // Межабзацные интервалы (перед и после)

            // Проверяем, что отступ слева либо отсутствует, либо равен нулю
            bool leftOk = indent?.Left == null || indent.Left.Value == "0";

            // Проверяем, что отступ справа либо отсутствует, либо равен нулю
            bool rightOk = indent?.Right == null || indent.Right.Value == "0";

            // Проверяем, что интервал перед абзацем либо отсутствует, либо равен нулю
            bool beforeOk = spacing?.Before == null || spacing.Before.Value == "0";

            // Проверяем, что интервал после абзаца либо отсутствует, либо равен нулю
            bool afterOk = spacing?.After == null || spacing.After.Value == "0";

            // Абзац соответствует требованиям, только если все 4 условия выполнены
            return leftOk && rightOk && beforeOk && afterOk;
        }
    }
}