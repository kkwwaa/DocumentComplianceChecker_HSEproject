using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace DocumentComplianceChecker_HSEproject.Services
{
    public class Exporter : IExporter
    {
        private readonly IFileManager _fileManager; // "Помощник" для работы с файлами

        // Конструктор: даём упаковщику помощника
        public Exporter(IFileManager fileManager)
        {
            _fileManager = fileManager; // Запоминаем помощника
        }

        // Упаковка документа
        public void ExportAnnotatedDocument(WordprocessingDocument sourceDoc, string outputPath)
        {
            #region
            //1.Проверяем входные данные
            if (sourceDoc == null) throw new ArgumentNullException("Документ не может быть null");
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("Неверный путь");

            try
            {
                // 2. Удаляем старую файл, если она есть
                if (_fileManager.FileExists(outputPath))
                    _fileManager.DeleteFile(outputPath);

                // 3. Создаём новый документ Word
                using (var newDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
                {
                    // 4. Копируем основное содержимое
                    var targetMainPart = newDoc.AddMainDocumentPart();
                    targetMainPart.Document = new Document(new Body(
                        sourceDoc.MainDocumentPart.Document.Body.CloneNode(true) // "Клонируем" содержимое
                    ));

                    // 5. Копируем стили, настройки
                    foreach (var part in sourceDoc.Parts)
                    {
                        if (part.OpenXmlPart != sourceDoc.MainDocumentPart)
                        {
                            newDoc.AddPart(part.OpenXmlPart, part.RelationshipId);
                        }
                    }
                } // Здесь using автоматически "закрывает коробку" (Dispose)
            }
            catch (Exception ex)
            {
                throw new Exception($"Не удалось упаковать документ: {ex.Message}");
            }
            #endregion
        }

        // Сохранение отчёта
        public void ExportReport(ValidationResult validationResult, string reportPath)
        {
            // Проверка на null
            if (validationResult == null || validationResult.Errors == null)
            {
                File.WriteAllText(reportPath, "Ошибки не найдены (null)");
                return;
            }

            // Формируем содержимое отчета
            var reportContent = $"Найдено ошибок: {validationResult.Errors.Count}\n" +
                              string.Join("\n", validationResult.Errors.Select(e =>
                                  $"- [{e.ErrorType}] {e.Message}\n" +
                                  $"  Текст: {e.ParagraphText?.Trim() ?? "не указан"}"));

            // Записываем в файл
            File.WriteAllText(reportPath, reportContent);
        }
    }
}