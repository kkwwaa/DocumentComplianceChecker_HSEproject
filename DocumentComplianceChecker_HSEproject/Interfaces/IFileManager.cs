namespace DocumentComplianceChecker_HSEproject.Interfaces
{
    public interface IFileManager
    {
        bool FileExists(string path);
        bool IsWordDocument(string path);
        string[] GetFilesFromDirectory(string directoryPath);
        string ReadAllText(string path);
        void WriteAllText(string path, string content);
        void DeleteFile(string path);
        void CopyFile(string sourcePath, string destinationPath);
        void WriteAllBytes(string path, byte[] data);
    }
}
