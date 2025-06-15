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
}