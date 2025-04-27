using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Rules;
using DocumentComplianceChecker_HSEproject.Services;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;

class Program
{
    // Группировка регистраций
    static void ConfigureServices(IServiceCollection services)
    {
        // Настройка DI
        services.AddSingleton<IFileManager, FileManager>(); // "Когда кто-то попросит IFileManager, верни FileManager"
        services.AddTransient<IDocumentLoader, DocumentLoader>();
        services.AddTransient<IExporter, Exporter>();
        services.AddTransient<IFormattingValidator, FormattingValidator>();
        services.AddTransient<AnnotationGenerator>();
    }

    static void Main()
    {
        // Создание TemplateManager и загрузка шаблона
        var templateManager = new TemplateManager();
        var template = templateManager.LoadTemplate("C:\\Users\\stepa\\Source\\Repos\\DocumentComplianceChecker_HSEproject\\DocumentComplianceChecker_HSEproject\\Templates\\default.docx");

        // Проверка на валидность шаблона
        if (template == null || !template.IsValid())
        {
            Console.WriteLine("Не удалось загрузить или шаблон невалиден.");
            return;
        }

        // Регистрируем правила, передавая в них шаблон
        var rules = new List<ValidationRule>
        {
            new FontRule(template),
            new FontSizeRule(template),
            new HeadingRule(template),
            new ColorRule(template)
        };

        // Регистрируем сервисы в DI контейнере
        var services = new ServiceCollection(); // Создаём коллекцию сервисов
        ConfigureServices(services); // Регистрируем зависимости
        services.AddSingleton(rules); // Регистрируем правила
        var provider = services.BuildServiceProvider(); // "Собираем" контейнер

        // Пример использования
        try
        {
            // DI делает это за вас:
            var fileManager = provider.GetRequiredService<IFileManager>(); // Грамотное создание
            var docLoader = provider.GetRequiredService<IDocumentLoader>();
            var exporter = provider.GetRequiredService<IExporter>();
            var validator = provider.GetRequiredService<IFormattingValidator>();
            var annotator = new AnnotationGenerator();

            string inputPath = "C:\\Users\\stepa\\OneDrive\\Рабочий стол\\input.docx";
            string outputPath = "C:\\Users\\stepa\\OneDrive\\Рабочий стол\\output.docx";
            string reportPath = "C:\\Users\\stepa\\OneDrive\\Рабочий стол\\report.txt";

            // Проверяем существование исходного файла
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
            exporter.ExportReport(errors, reportPath); // Сохраняем отчёт

            Console.WriteLine($"Проверка завершена. Результаты сохранены в:\n{outputPath}\n{reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка: {ex.Message}");
        }
    }
}