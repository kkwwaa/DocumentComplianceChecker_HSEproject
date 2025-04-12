using DocumentFormat.OpenXml.Packaging;

namespace DocumentComplianceChecker_HSEproject.Interfaces
{
    // Disposable Явно показывает, что реализация требует очистки
    //Обязывает все реализации включать Dispose
    public interface IDocumentLoader : IDisposable 
    {
        // Метод для открытия документа
        WordprocessingDocument LoadDocument(string path);

        // Метод для сохранения документа
        void SaveDocument(WordprocessingDocument doc);
    }
}
