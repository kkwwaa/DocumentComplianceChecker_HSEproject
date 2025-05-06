using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Rules;
using DocumentComplianceChecker_HSEproject.Services;
using DocumentFormat.OpenXml.Office2021.Excel.NamedSheetViews;
using Microsoft.Extensions.DependencyInjection;

static void ConfigureServices(IServiceCollection services)
{
    // Регистрируем все зависимости в DI контейнере
    services.AddSingleton<IFileManager, FileManager>(); // FileManager будет единственным экземпляром
    services.AddTransient<IDocumentLoader, DocumentLoader>(); // DocumentLoader будет создаваться каждый раз
    services.AddTransient<IExporter, Exporter>(); // Exporter тоже будет создаваться каждый раз
    services.AddTransient<IFormattingValidator, FormattingValidator>(); // FormattingValidator каждый раз
    services.AddTransient<AnnotationGenerator>(); // AnnotationGenerator каждый раз
}

// Создаём DI контейнер
var services = new ServiceCollection();
ConfigureServices(services); // Конфигурируем контейнер с необходимыми зависимостями

// Меню выбора
Console.WriteLine("Выберите режим проверки:");
Console.WriteLine("1. Использовать стандартные правила");
Console.WriteLine("2. Использовать шаблон (Template)");
Console.Write("Введите номер варианта: ");
string choice = Console.ReadLine();

List<ValidationRule> rules;
bool useTemplate = false;

if (choice == "1")
{
    // Если выбраны стандартные правила, создаем и добавляем их в список
    rules = new List<ValidationRule>
    {
        new ParagraphStyleAndSizeRule
        {
            ErrorMessage = "Стиль или формат заголовка/текста не соответствует требованиям"
        },
        new ColorRule
        {
            AllowedColors = new List<string> { "auto", "000000" }, // Цвет шрифта должен быть чёрным или авто
            ErrorMessage = "Цвет шрифта должен быть чёрным или авто"
        },
        new PageMarginRule
        {
            ErrorMessage = "Поля документа не соответствуют требованиям (верх: 2см, низ: 2см, лево: 3см, право: 1.5см)"
        },
        new LineSpacingRule
        {
            RequiredLineSpacing = 1.5, // Устанавливаем обязательный межстрочный интервал
            ErrorMessage = "Межстрочный интервал должен быть 1,5"
        },
        new FirstLineIndentRule
        {
            RequiredIndentInCm = 1.25, // Красная строка должна быть 1,25 см
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

    // Добавляем список правил в DI контейнер
    services.AddSingleton(rules);
}
else if (choice == "2")
{
    // Если выбран шаблон, указываем, что будем использовать его
    useTemplate = true;
    var template = new Template(); // Создаём объект шаблона
    services.AddSingleton(template); // Добавляем шаблон в DI контейнер
}
else
{
    // Если пользователь ввёл неверный выбор, завершаем программу
    Console.WriteLine("Неверный выбор. Завершение программы.");
    return;
}

// Строим провайдер для получения зависимостей
var provider = services.BuildServiceProvider();

try
{
    // Получаем необходимые сервисы из DI контейнера
    var fileManager = provider.GetRequiredService<IFileManager>();
    var docLoader = provider.GetRequiredService<IDocumentLoader>();
    var exporter = provider.GetRequiredService<IExporter>();
    var annotator = provider.GetRequiredService<AnnotationGenerator>();

    // Пути к входному и выходному файлам, а также к отчету
    string inputPath = "C:\\Users\\stepa\\OneDrive\\Рабочий стол\\input.docx";
    string outputPath = "C:\\Users\\stepa\\OneDrive\\Рабочий стол\\output.docx";
    string reportPath = "C:\\Users\\stepa\\OneDrive\\Рабочий стол\\report.txt";

    // Проверка существования файла
    if (!fileManager.FileExists(inputPath))
    {
        Console.WriteLine($"Файл {inputPath} не найден!");
        return;
    }

    // Создаем копию документа для обработки
    using var doc = docLoader.CreateDocumentCopy(inputPath, outputPath);

    List<Error> errors;

    // Если выбран шаблон, используем его для валидации
    if (useTemplate)
    {
        var template = provider.GetRequiredService<Template>(); // Получаем шаблон из DI
        var formatter = new FormattingTemplate(template); // Создаем объект для валидации с использованием шаблона
        errors = new List<Error>();

        // Обрабатываем каждый параграф
        foreach (var paragraph in doc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
        {
            var result = formatter.ValidateParagraph(paragraph); // Получаем ошибки для параграфа
            errors.AddRange(result.Errors); // Добавляем ошибки в общий список
        }
    }
    else
    {
        // Если шаблон не выбран, используем стандартную валидацию
        var validator = provider.GetRequiredService<IFormattingValidator>();
        var formattingValidationResult = validator.Validate(doc);
        errors = formattingValidationResult.Errors; // Просто присваиваем ошибки в переменную
    }

    // Преобразуем List<Error> в ValidationResult
    var validationResult = new ValidationResult();
    validationResult.Errors.AddRange(errors);

    // Применяем аннотации к документу на основе ошибок
    annotator.ApplyAnnotations(doc, validationResult);
    // Экспортируем отчет
    exporter.ExportReport(validationResult, reportPath);

    // Выводим сообщение о завершении работы программы
    Console.WriteLine($"Проверка завершена. Результаты сохранены в:\n{outputPath}\n{reportPath}");
}
catch (Exception ex)
{
    // Если произошла ошибка, выводим сообщение
    Console.WriteLine($"Ошибка: {ex.Message}");
}