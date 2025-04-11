namespace DocumentComplianceChecker_HSEproject.Models
{
    public class FormattingError
    {
        public string ErrorType { get; set; }
        public string Message { get; set; }
        public string Location { get; set; } = "Unknown";
    }
}
