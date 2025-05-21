using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Interfaces
{
    // Интерфейс для проверки свойств параграфа (выравнивание, отступы и т.д.)
    public interface IParagraphValidationRule
    {
        string ErrorMessage { get; }
        bool ValidateParagraph(Paragraph paragraph);
    }

    // Интерфейс для проверки свойств Run (шрифт, размер и т.д.)
    public interface IRunValidationRule
    {
        string ErrorMessage { get; }
        bool ValidateRun(Paragraph paragraph, Run run);
    }
}