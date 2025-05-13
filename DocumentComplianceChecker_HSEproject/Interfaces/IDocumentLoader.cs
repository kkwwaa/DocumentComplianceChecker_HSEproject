using DocumentFormat.OpenXml.Packaging;

namespace DocumentComplianceChecker_HSEproject.Interfaces
{
    // Disposable Явно показывает, что реализация требует очистки
    //Обязывает все реализации включать Dispose
    public interface IDocumentLoader : IDisposable 
    {
        // Метод для клонирования и открытия
        WordprocessingDocument CreateDocumentCopy(string inputPath, string outputPath);

        // Метод для сохранения документа
        void SaveDocument(WordprocessingDocument doc);
    }
}