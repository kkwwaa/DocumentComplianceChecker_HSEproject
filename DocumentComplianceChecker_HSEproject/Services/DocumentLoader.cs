﻿using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentComplianceChecker_HSEproject
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

        // Открываем документ Word
        public WordprocessingDocument LoadDocument(string path)
        {
            // Проверяем через помощника - это действительно Word-файл?
            if (!_fileManager.IsWordDocument(path))
                throw new ArgumentException("File is not a Word document (.docx)");

            // Открываем документ для редактирования (true = можно изменять)
            _currentDocument = WordprocessingDocument.Open(path, true);
            // Возвращаем открытый документ
            return _currentDocument;
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
