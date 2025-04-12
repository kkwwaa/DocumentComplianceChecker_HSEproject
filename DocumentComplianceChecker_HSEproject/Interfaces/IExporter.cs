using DocumentFormat.OpenXml.Packaging;

namespace DocumentComplianceChecker_HSEproject.Interfaces
{
    public interface IExporter
    {
        // копирует Word-документ с аннотациями (документ -> новый файл)
        void ExportAnnotatedDocument(WordprocessingDocument doc, string outputPath);

        // сохраняет текстовый отчет (текст -> файл отчёта)
        void ExportReport(string reportContent, string outputPath);

    }
}
