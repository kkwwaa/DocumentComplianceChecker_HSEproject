using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentComplianceChecker_HSEproject.Interfaces
{
    /// <summary>
    /// Интерфейс валидатора форматирования документов
    /// </summary>
    public interface IFormattingValidator
    {
        /// <summary>
        /// Выполняет проверку документа по заданным правилам
        /// </summary>
        /// <param name="doc">Открытый Word-документ</param>
        /// <returns>Результат валидации с коллекцией ошибок</returns>
        ValidationResult Validate(WordprocessingDocument doc);
    }

    /// <summary>
    /// Интерфейс фабрики для создания валидатора
    /// </summary>
    public interface IFormattingValidatorFactory
    {
        /// <summary>
        /// Создает экземпляр валидатора с указанными правилами
        /// </summary>
        /// <param name="rules">Коллекция правил проверки</param>
        IFormattingValidator CreateValidator(IEnumerable<ValidationRule> rules);
    }
}
