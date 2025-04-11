using DocumentFormat.OpenXml.Packaging;

namespace DocumentComplianceChecker_HSEproject.Interfaces
{
    public interface IDocumentLoader : IDisposable
    {
        WordprocessingDocument LoadDocument(string path);
        void SaveDocument(WordprocessingDocument doc);
    }
}
