namespace DocumentComplianceChecker_HSEproject.Models
{
    public class ValidationResult
    {
        public List<Error> Errors { get; } = new List<Error>();
    }
}