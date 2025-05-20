using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Interfaces
{
    public interface IValidationRule
    {
        string ErrorMessage { get; }
        bool RuleValidator(Paragraph paragraph, Run run = null);
    }
}