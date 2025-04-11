using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentComplianceChecker_HSEproject
{
    public class FormattingValidator : IFormattingValidator
    {
        public List<FormattingError> Validate(WordprocessingDocument doc)
        {
            // Заглушка - возвращаем тестовые ошибки
            return new List<FormattingError>
        {
            new FormattingError { ErrorType = "Font", Message = "Шрифт должен быть Times New Roman" },
            new FormattingError { ErrorType = "Spacing", Message = "Межстрочный интервал должен быть 1.5" }
        };
        }
    }
}
