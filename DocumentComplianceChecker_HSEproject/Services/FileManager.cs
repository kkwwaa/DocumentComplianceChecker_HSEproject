using DocumentComplianceChecker_HSEproject.Interfaces;

namespace DocumentComplianceChecker_HSEproject
{
    public class FileManager : IFileManager
    {
        // => означает "возвращает"
        // Компилятор автоматически подставляет return
        // Проверяет существует ли файл

        #region Текстовые операции
        // Читает текст из файла
        public string ReadAllText(string path) 
            => File.ReadAllText(path);

        // Записывает текст в файл
        public void WriteAllText(string path, string content) 
            => File.WriteAllText(path, content);
        #endregion

        #region Бинарные операции
        public byte[] ReadAllBytes(string path) => File.ReadAllBytes(path);

        // Сохраняете изменённый Word-документ
        // Копируете файлы
        // Работаете с шаблонами документов
        // Собирает докс
        public void WriteAllBytes(string path, byte[] data)
            => File.WriteAllBytes(path, data);
        #endregion

        #region Файловые операции
        public bool FileExists(string path) => File.Exists(path);

        // Проверяет .docx ли это (независимо от регистра .DOCX/.docx)
        public bool IsWordDocument(string path)
            => Path.GetExtension(path).Equals(".docx", StringComparison.OrdinalIgnoreCase);

        // Получает все .docx файлы из папки
        public string[] GetFilesFromDirectory(string directoryPath)
            => Directory.GetFiles(directoryPath, "*.docx");

        // Удаление файла по указанному пути
        public void DeleteFile(string path) => File.Delete(path); // fileManager.DeleteFile("C:/docs/old.docx"); // Безвозвратное удаление

        // Копирование файла
        public void CopyFile(string sourcePath, string destinationPath)
            => File.Copy(sourcePath, destinationPath, true); // fileManager.CopyFile("C:/docs/source.docx", "D:/backup/copy.docx");
        #endregion

    }
}
