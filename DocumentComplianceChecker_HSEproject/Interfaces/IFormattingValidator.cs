using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentComplianceChecker_HSEproject.Interfaces
{
    public interface IFormattingValidator
    {
        List<FormattingError> Validate(WordprocessingDocument doc);
    }
}
