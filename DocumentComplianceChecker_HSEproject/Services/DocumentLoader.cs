using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentComplianceChecker_HSEproject.Services
{
    public class DocumentLoader : IDocumentLoader
    {
        // Здесь храним открытый документ
        private WordprocessingDocument? _currentDocument;
        // "Помощник" для работы с файлами
        private readonly IFileManager _fileManager;

        // Конструктор (получаем помощника при создании)
        public DocumentLoader(IFileManager fileManager)
        {
            // Запоминаем помощника для работы с файлами
            _fileManager = fileManager;
        }
       
        // Копируем файл
        public WordprocessingDocument CreateDocumentCopy(string inputPath, string outputPath)
        {
            // Проверяем через помощника - это действительно Word-файл?
            if (!_fileManager.IsWordDocument(inputPath))
                throw new ArgumentException("File is not a Word document (.docx)");

            // Копируем файл
            File.Copy(inputPath, outputPath, overwrite: true);

            // Открываем копию с возможностью редактирования
            return WordprocessingDocument.Open(outputPath, true);
        }

        // Сохраняем изменения в документе
        public void SaveDocument(WordprocessingDocument doc)
        {
            // Сохраняем все изменения
            doc.Save();
            // Сбрасываем ссылку на текущий документ
            _currentDocument = null;
        }

        // Метод для уборки
        public void Dispose()
        {
            // Закрываем документ, если он открыт
            _currentDocument?.Dispose();
            // Говорим системе, что уборку сделали
            GC.SuppressFinalize(this);
        }
    }
}