namespace DocumentComplianceChecker_HSEproject.Models
{
    // Models/Error.cs
    public class Error
    {
        public string ErrorType { get; set; }
        public string Message { get; set; }
        public string ParagraphText { get; set; } // Текст абзаца для аннотации
    }
}
