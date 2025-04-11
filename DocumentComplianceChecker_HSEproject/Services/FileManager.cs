using DocumentComplianceChecker_HSEproject.Interfaces;

namespace DocumentComplianceChecker_HSEproject
{
    public class FileManager : IFileManager
    {
        public bool FileExists(string path) => File.Exists(path);

        public bool IsWordDocument(string path)
            => Path.GetExtension(path).Equals(".docx", StringComparison.OrdinalIgnoreCase);

        public string[] GetFilesFromDirectory(string directoryPath)
            => Directory.GetFiles(directoryPath, "*.docx");

        public string ReadAllText(string path) => File.ReadAllText(path);

        public void WriteAllText(string path, string content)
            => File.WriteAllText(path, content);

        public void DeleteFile(string path) => File.Delete(path);

        public void CopyFile(string sourcePath, string destinationPath)
            => File.Copy(sourcePath, destinationPath, true);

        public void WriteAllBytes(string path, byte[] data)
            => File.WriteAllBytes(path, data);
    }
}
