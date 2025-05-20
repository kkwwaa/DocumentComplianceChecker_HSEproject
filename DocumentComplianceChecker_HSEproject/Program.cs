using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Services;
using Microsoft.Extensions.DependencyInjection;

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
    new BasicRules.NormalStyleRule(),
    new BasicRules.Heading1Rule(),
    new BasicRules.Heading2Rule(),
    new BasicRules.Heading3Rule(),
    new BasicRules.PageMarginRule()
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

    HashSet<string> uniqueErrors = new HashSet<string>();

    if (useTemplate)
    {
        var template = provider.GetRequiredService<Template>();
        var formatter = new FormattingTemplate(template);
        validationResult = formatter.ValidateDocument(doc);

        // Проходим по всем ошибкам
        foreach (var error in validationResult.Errors)
        {
            uniqueErrors.Add(error.Message); // Собираем уникальные ошибки
        }
    }
    else
    {
        var validator = provider.GetRequiredService<IFormattingValidator>();
        validationResult = validator.Validate(doc);

        // Проходим по всем ошибкам
        foreach (var error in validationResult.Errors)
        {
            uniqueErrors.Add(error.Message); // Собираем уникальные ошибки
        }
    }

    // Применяем аннотации на основе уникальных ошибок
    annotator.ApplyAnnotations(doc, validationResult);

    // Создаём отчет
    exporter.ExportReport(validationResult, reportPath);

    // Вывод всех ошибок
    Console.WriteLine("Ошибки по абзацам:");
    foreach (var error in uniqueErrors)
    {
        Console.WriteLine(error); // Выводим уникальные ошибки
    }

    // Выводим путь к результатам
    Console.WriteLine($"Проверка завершена. Результаты сохранены в:\n{outputPath}\n{reportPath}");
}
catch (Exception ex)
{
    Console.WriteLine($"Ошибка: {ex.Message}");
}