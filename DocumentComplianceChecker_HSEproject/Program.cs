using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Rules;
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

List<ValidationRule> rules = null;
bool useTemplate = false;

if (choice == "1")
{
    rules = new List<ValidationRule>
    {
        new ParagraphStyleAndSizeRule
        {
            ErrorMessage = "Стиль или формат заголовка/текста не соответствует требованиям"
        },
        new ColorRule
        {
            AllowedColors = new List<string> { "auto", "000000" },
            ErrorMessage = "Цвет шрифта должен быть чёрным или авто"
        },
        new PageMarginRule
        {
            ErrorMessage = "Поля документа не соответствуют требованиям (верх: 2см, низ: 2см, лево: 3см, право: 1.5см)"
        },
        new LineSpacingRule
        {
            RequiredLineSpacing = 1.5,
            ErrorMessage = "Межстрочный интервал должен быть 1,5"
        },
        new FirstLineIndentRule
        {
            RequiredIndentInCm = 1.25,
            ErrorMessage = "Красная строка должна быть 1,25 см"
        },
        new JustificationRule
        {
            ErrorMessage = "Абзац должен быть выровнен по ширине"
        },
        new ParagraphSpacingRule
        {
            ErrorMessage = "Абзацы должны иметь отступы: слева = 0, справа = 0, перед = 0, после = 0"
        }
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

    if (useTemplate)
    {
        var template = provider.GetRequiredService<Template>();
        var formatter = new FormattingTemplate(template);
        validationResult = formatter.ValidateDocument(doc); // Новый метод
    }
    else
    {
        var validator = provider.GetRequiredService<IFormattingValidator>();
        validationResult = validator.Validate(doc); // Метод возвращает ValidationResult
    }

    annotator.ApplyAnnotations(doc, validationResult);
    exporter.ExportReport(validationResult, reportPath);

    Console.WriteLine($"Проверка завершена. Результаты сохранены в:\n{outputPath}\n{reportPath}");
}
catch (Exception ex)
{
    Console.WriteLine($"Ошибка: {ex.Message}");
}
