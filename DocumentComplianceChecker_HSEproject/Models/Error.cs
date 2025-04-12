using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Models
{
    // Models/Error.cs
    public class Error
    {
        public string ErrorType { get; set; }
        public string Message { get; set; }
        public string ParagraphText { get; set; }
        public int ParagraphIndex { get; set; } // Новое поле
        public Run TargetRun { get; set; }
    }
}
