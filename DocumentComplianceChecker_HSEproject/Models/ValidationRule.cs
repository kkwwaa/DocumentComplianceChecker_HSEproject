using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Models
{
    public interface IValidationRule
    {
        string ErrorMessage { get; set; }
        bool RuleValidator(Paragraph paragraph, Run run = null);
    }
}