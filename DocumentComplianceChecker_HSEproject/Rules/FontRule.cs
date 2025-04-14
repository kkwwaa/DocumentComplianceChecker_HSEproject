using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Services;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    public class FontRule : ValidationRule
    {
        private readonly FormattingTemplate _template;

        public FontRule(FormattingTemplate template)
        {
            _template = template;
        }

        public override bool Validate(Paragraph paragraph, Run run = null)
        {
            throw new NotImplementedException();
        }
    }
}
