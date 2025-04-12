using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Services;
using Microsoft.Extensions.DependencyInjection;

// Группировка регистраций
//AddSingleton – один экземпляр на всю программу (например, FileManager).
//AddTransient – новый экземпляр при каждом запросе (например, DocumentLoader).
static void ConfigureServices(IServiceCollection services)
{// Настройка DI    
    services.AddSingleton<IFileManager, FileManager>();// "Когда кто-то попросит IFileManager, верни FileManager"
    services.AddTransient<IDocumentLoader, DocumentLoader>();
    services.AddTransient<IExporter, Exporter>();
    services.AddTransient<IFormattingValidator, FormattingValidator>();    
}

var services = new ServiceCollection(); // Создаём коллекцию сервисов
ConfigureServices(services); // Регистрируем зависимости
var provider = services.BuildServiceProvider(); // "Собираем" контейнер

// Пример использования
try
{
    // DI делает это за вас:
    var fileManager = provider.GetRequiredService<IFileManager>();// Грамотное создание
    // (возвращает FileManager, но вы работаете только с интерфейсом)  
    var docLoader = provider.GetRequiredService<IDocumentLoader>();
    var exporter = provider.GetRequiredService<IExporter>();
    var validator = provider.GetRequiredService<IFormattingValidator>();
    var annotator = new AnnotationGenerator();

    string inputPath = "input.docx";
    string outputPath = "output.docx";
    string reportPath = "report.txt";
    

    if (!fileManager.FileExists(inputPath))
    {
        Console.WriteLine($"Файл {inputPath} не найден!");
        return;
    }

    // Основной workflow
    using var doc = docLoader.LoadDocument("input.docx");

    // Проверка форматирования
    var errors = validator.Validate(doc);

    // Аннотирование ошибок
    annotator.ApplyAnnotations(doc, errors);

    // Сохраняем результаты
    exporter.ExportAnnotatedDocument(doc, outputPath);  // Сохраняем исправленный документ
    exporter.ExportReport(errors, reportPath);   // Сохраняем отчёт

    Console.WriteLine($"Проверка завершена. Результаты сохранены в:\n{outputPath}\n{reportPath}");
}
catch (Exception ex)
{
    Console.WriteLine($"Ошибка: {ex.Message}");
}