using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    internal class HeadingError : Error
    {
        public string Message { get; set; }
        public Paragraph Paragraph { get; set; }
        public object AnnotationType { get; set; }
    }
}