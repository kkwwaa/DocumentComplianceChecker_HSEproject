using DocumentComplianceChecker_HSEproject.Interfaces; // Подключаем пространство имён с интерфейсом IValidationRule
using DocumentFormat.OpenXml.Wordprocessing; // Подключаем библиотеку OpenXML для работы с Word-документами

// Абстрактный базовый класс для правил валидации, реализующий оба интерфейса
internal abstract class BasicRules : IParagraphValidationRule, IRunValidationRule
{
    // Абстрактное свойство для сообщения об ошибке, реализуется в дочерних классах
    public abstract string ErrorMessage { get; }

    // Абстрактный метод для проверки свойств параграфа
    public abstract bool ValidateParagraph(Paragraph paragraph);

    // Абстрактный метод для проверки свойств Run
    public abstract bool ValidateRun(Paragraph paragraph, Run run);

    // Получает ID стиля параграфа, возвращает "Normal", если стиль не указан
    protected string GetStyleId(Paragraph paragraph)
    {
        return paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal";
    }
}