using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Models
{
    public abstract class ValidationRule
    {
        public string ErrorMessage { get; set; }
        public abstract bool Validate(Paragraph paragraph, Run run = null);

    }
}
