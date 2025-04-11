using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentComplianceChecker_HSEproject
{
    public class DocumentLoader : IDocumentLoader
    {
        private WordprocessingDocument? _currentDocument;
        private readonly IFileManager _fileManager;

        public DocumentLoader(IFileManager fileManager)
        {
            _fileManager = fileManager;
        }

        public WordprocessingDocument LoadDocument(string path)
        {
            if (!_fileManager.IsWordDocument(path))
                throw new ArgumentException("File is not a Word document (.docx)");

            _currentDocument = WordprocessingDocument.Open(path, true);
            return _currentDocument;
        }

        public void SaveDocument(WordprocessingDocument doc)
        {
            doc.Save();
            _currentDocument = null;
        }

        public void Dispose()
        {
            _currentDocument?.Dispose();
            GC.SuppressFinalize(this);
        }
    }
}
