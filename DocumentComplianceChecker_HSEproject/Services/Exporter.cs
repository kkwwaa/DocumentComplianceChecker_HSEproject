using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Services
{
    public class Exporter : IExporter
    {
        private readonly IFileManager _fileManager;

        public Exporter(IFileManager fileManager)
        {
            _fileManager = fileManager;
        }

        public void ExportAnnotatedDocument(WordprocessingDocument sourceDoc, string outputPath)
        {
            try
            {
                // 1. Удаляем существующий файл, если есть
                if (_fileManager.FileExists(outputPath))
                {
                    _fileManager.DeleteFile(outputPath);
                }

                // 2. Создаем новый документ
                using (var newDoc = WordprocessingDocument.Create(
                    outputPath,
                    WordprocessingDocumentType.Document))
                {
                    // 3. Копируем основные части документа
                    CopyMainDocumentPart(sourceDoc, newDoc);

                    // 4. Копируем все остальные части (стили, настройки и т.д.)
                    foreach (var part in sourceDoc.Parts)
                    {
                        if (part.OpenXmlPart != sourceDoc.MainDocumentPart)
                        {
                            newDoc.AddPart(part.OpenXmlPart, part.RelationshipId);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Export failed: {ex.Message}", ex);
            }
        }

        private void CopyMainDocumentPart(WordprocessingDocument source, WordprocessingDocument target)
        {
            // Копируем содержимое основного документа
            var sourceBody = source.MainDocumentPart.Document.Body;
            var targetMainPart = target.AddMainDocumentPart();

            targetMainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new Body(sourceBody.OuterXml));
        }

        public void ExportReport(string reportContent, string outputPath)
        {
            _fileManager.WriteAllText(outputPath, reportContent);
        }
    }
}
