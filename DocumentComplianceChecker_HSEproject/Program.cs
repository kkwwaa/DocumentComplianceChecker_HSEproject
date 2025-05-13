using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Services;
using Microsoft.Extensions.DependencyInjection;
using System.Text.RegularExpressions;

static void ConfigureServices(IServiceCollection services)
{
    services.AddSingleton<IFileManager, FileManager>();
    services.AddTransient<IDocumentLoader, DocumentLoader>();
    services.AddTransient<IExporter, Exporter>();
    services.AddTransient<IFormattingValidator, FormattingValidator>();
    services.AddTransient<AnnotationGenerator>();
}

// Создаём DI контейнер
var services = new ServiceCollection();
ConfigureServices(services);

// Меню выбора
Console.WriteLine("Выберите режим проверки:");
Console.WriteLine("1. Использовать стандартные правила");
Console.WriteLine("2. Использовать шаблон (Template)");
Console.Write("Введите номер варианта: ");
string choice = Console.ReadLine();

List<IValidationRule> rules = null;
bool useTemplate = false;

if (choice == "1")
{
    rules = new List<IValidationRule>
    {
        new BasicRules.ColorRule(),
        new BasicRules.FirstLineIndentRule(),
        new BasicRules.JustificationRule(),
        new BasicRules.LineSpacingRule(),
        new BasicRules.PageMarginRule(),
        new BasicRules.ParagraphSpacingRule(),
        new BasicRules.ParagraphStyleAndSizeRule(),
        new BasicRules.HeadingStartsNewPageRule(),
        new BasicRules.HeadingSpacingRule(),
        new BasicRules.Heading3NotInTOCRule()
    };

    services.AddSingleton(rules);
}
else if (choice == "2")
{
    useTemplate = true;
    var template = new Template();
    services.AddSingleton(template);
}
else
{
    Console.WriteLine("Неверный выбор. Завершение программы.");
    return;
}

var provider = services.BuildServiceProvider();

try
{
    var fileManager = provider.GetRequiredService<IFileManager>();
    var docLoader = provider.GetRequiredService<IDocumentLoader>();
    var exporter = provider.GetRequiredService<IExporter>();
    var annotator = provider.GetRequiredService<AnnotationGenerator>();

    string inputPath = "C:\\Users\\stepa\\OneDrive\\Рабочий стол\\input.docx";
    string outputPath = "C:\\Users\\stepa\\OneDrive\\Рабочий стол\\output.docx";
    string reportPath = "C:\\Users\\stepa\\OneDrive\\Рабочий стол\\report.txt";

    if (!fileManager.FileExists(inputPath))
    {
        Console.WriteLine($"Файл {inputPath} не найден!");
        return;
    }

    using var doc = docLoader.CreateDocumentCopy(inputPath, outputPath);

    // Итоговый объект с результатами валидации
    ValidationResult validationResult;
    List<string> errorMessages = new List<string>();

    if (useTemplate)
    {
        var template = provider.GetRequiredService<Template>();
        var formatter = new FormattingTemplate(template);
        validationResult = formatter.ValidateDocument(doc); // Новый метод

        // Собираем все ошибки в список
        foreach (var error in validationResult.Errors)
        {
            // Простой способ получения строки из ошибки, если она не связана с позицией
            errorMessages.Add($"Ошибка: {error.Message}");
        }
    }
    else
    {
        var validator = provider.GetRequiredService<IFormattingValidator>();
        validationResult = validator.Validate(doc); // Метод возвращает ValidationResult

        // Собираем все ошибки в список
        foreach (var error in validationResult.Errors)
        {
            // Простой способ получения строки из ошибки, если она не связана с позицией
            errorMessages.Add($"Ошибка: {error.Message}");
        }
    }

    annotator.ApplyAnnotations(doc, validationResult);
    exporter.ExportReport(validationResult, reportPath);

    // Вывод всех ошибок
    Console.WriteLine("Ошибки по строкам:");
    foreach (var errorMessage in errorMessages)
    {
        Console.WriteLine(errorMessage);
    }

    Console.WriteLine($"Проверка завершена. Результаты сохранены в:\n{outputPath}\n{reportPath}");
}
catch (Exception ex)
{
    Console.WriteLine($"Ошибка: {ex.Message}");
}
