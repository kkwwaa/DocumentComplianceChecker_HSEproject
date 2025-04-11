using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Services;
using DocumentComplianceChecker_HSEproject;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.DependencyInjection;

// Настройка DI
var services = new ServiceCollection();
services.AddSingleton<IFileManager, FileManager>();
services.AddTransient<IDocumentLoader, DocumentLoader>();
services.AddTransient<IExporter, Exporter>();
services.AddTransient<IFormattingValidator, FormattingValidator>();

var provider = services.BuildServiceProvider();

// Пример использования
try
{
    var fileManager = provider.GetRequiredService<IFileManager>();
    var docLoader = provider.GetRequiredService<IDocumentLoader>();
    var exporter = provider.GetRequiredService<IExporter>();
    var validator = provider.GetRequiredService<IFormattingValidator>();

    string inputPath = "input.docx";
    string outputPath = "output.docx";
    string reportPath = "report.txt";

    if (!fileManager.FileExists(inputPath))
    {
        Console.WriteLine($"Файл {inputPath} не найден!");
        return;
    }

    // Основной workflow
    using var doc = docLoader.LoadDocument(inputPath);
    var errors = validator.Validate(doc);

    // Генерируем простой отчет
    string reportContent = $"Найдено ошибок: {errors.Count}\n" +
                          string.Join("\n", errors.Select(e => $"- {e.ErrorType}: {e.Message}"));

    // Сохраняем результаты
    exporter.ExportAnnotatedDocument(doc, outputPath);
    exporter.ExportReport(reportContent, reportPath);

    Console.WriteLine($"Проверка завершена. Результаты сохранены в:\n{outputPath}\n{reportPath}");
}
catch (Exception ex)
{
    Console.WriteLine($"Ошибка: {ex.Message}");
}