using DocumentFormat.OpenXml.Packaging;

namespace DocumentComplianceChecker_HSEproject.Interfaces
{
    public interface IDocumentLoader : IDisposable
    {
        // Метод для открытия документа
        WordprocessingDocument LoadDocument(string path);

        // Метод для сохранения документа
        void SaveDocument(WordprocessingDocument doc);
    }
}
