﻿using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Rules;
using DocumentComplianceChecker_HSEproject.Services;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
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
    services.AddTransient<AnnotationGenerator>();
}

var templateManager = new TemplateManager();
var template = templateManager.LoadTemplate("C:\\Users\\stepa\\Source\\Repos\\DocumentComplianceChecker_HSEproject\\DocumentComplianceChecker_HSEproject\\Templates\\default.json");

// передаёшь template в каждое правило
var rules = new List<ValidationRule>
{
    new FontRule(template),
    new FontSizeRule(template),
    new HeadingRule(template),
    new ColorRule(template)
};


var services = new ServiceCollection(); // Создаём коллекцию сервисов
ConfigureServices(services); // Регистрируем зависимости
services.AddSingleton(rules); // Регистрируем правила
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

    string inputPath = "C:\\Users\\stepa\\OneDrive\\Рабочий стол\\input.docx";
    string outputPath = "C:\\Users\\stepa\\OneDrive\\Рабочий стол\\output.docx";
    string reportPath = "C:\\Users\\stepa\\OneDrive\\Рабочий стол\\report.txt";
    

    if (!fileManager.FileExists(inputPath))
    {
        Console.WriteLine($"Файл {inputPath} не найден!");
        return;
    }

    // Основной workflow
    using var doc = docLoader.CreateDocumentCopy(inputPath, outputPath);

    // Проверка форматирования
    var errors = validator.Validate(doc);

    // Аннотирование ошибок
    annotator.ApplyAnnotations(doc, errors);

    // Сохраняем результаты
    exporter.ExportReport(errors, reportPath);   // Сохраняем отчёт

    Console.WriteLine($"Проверка завершена. Результаты сохранены в:\n{outputPath}\n{reportPath}");
}
catch (Exception ex)
{
    Console.WriteLine($"Ошибка: {ex.Message}");
}