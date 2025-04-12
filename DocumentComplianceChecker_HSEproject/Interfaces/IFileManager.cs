namespace DocumentComplianceChecker_HSEproject.Interfaces
{
    public interface IFileManager
    {
        //"Любой класс, который реализует IFileManager, ДОЛЖЕН иметь все эти методы"
        // Интерфейс для работы с файловой системой
        // ================= Текстовые операции =================

        // Читает весь текст из файла
        string ReadAllText(string path);

        // Записывает текст в файл (перезаписывает, если файл существует)
        void WriteAllText(string path, string content);

        // ================= Бинарные операции =================

        // Читает файл как массив байтов
        byte[] ReadAllBytes(string path);

        // Записывает массив байтов в файл
        void WriteAllBytes(string path, byte[] data);

        // ================= Файловые операции =================

        // Проверяет существование файла
        bool FileExists(string path);

        // Удаляет файл
        void DeleteFile(string path);

        // Копирует файл с возможностью перезаписи
        void CopyFile(string sourcePath, string destinationPath);

        // Проверяет, является ли файл документом Word (.docx)
        bool IsWordDocument(string path);
    }
}
