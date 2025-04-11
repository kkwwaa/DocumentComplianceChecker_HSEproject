using DocumentFormat.OpenXml.Packaging;

namespace DocumentComplianceChecker_HSEproject.Interfaces
{
    public interface IExporter
    {
        void ExportAnnotatedDocument(WordprocessingDocument doc, string outputPath);
        void ExportReport(string reportContent, string outputPath);
    }
}
