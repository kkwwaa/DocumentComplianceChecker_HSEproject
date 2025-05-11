using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Rules;
using DocumentComplianceChecker_HSEproject.Services;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using System.Windows;
using System.Windows.Controls;

namespace DocumentComplianceChecker_HSEproject
{
    using MyValidationRule = DocumentComplianceChecker_HSEproject.Models.ValidationRule;
    using MyValidationResult = DocumentComplianceChecker_HSEproject.Models.ValidationResult;

    public partial class MainWindow : Window
    {
        private readonly IServiceProvider provider;
        private List<MyValidationRule> _rules;
        private bool _useTemplate;

        public MainWindow()
        {
            InitializeComponent();

          
        }

        private void ConfigureServices(IServiceCollection services)
        {
            services.AddSingleton<IFileManager, FileManager>();
            services.AddTransient<IDocumentLoader, DocumentLoader>();
            services.AddTransient<IExporter, Exporter>();
            services.AddTransient<IFormattingValidator, FormattingValidator>();
            services.AddTransient<AnnotationGenerator>();
        }

        private void CheckDocument_Click(object sender, RoutedEventArgs e)
        {
            // Настройка DI
            var services = new ServiceCollection();
            ConfigureServices(services); // Конфигурируем контейнер с необходимыми зависимостями

            try
            {
                // Определяем режим проверки на основе выбора пользователя
                _useTemplate = TemplateRadio.IsChecked == true;
                List<MyValidationRule> _rules;

                if (_useTemplate)
                {
                    var template = new Template(); // Создаём объект шаблона
                    services.AddSingleton(template); // Добавляем шаблон в DI контейнер
                }
                else
                {
                    _rules = new List<MyValidationRule>
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
                    // Добавляем список правил в DI контейнер
                    services.AddSingleton(_rules);
                    //provider.GetService<IServiceCollection>()?.AddSingleton(_rules);
                }


                // Строим провайдер для получения зависимостей
                var provider = services.BuildServiceProvider();

                // Получаем сервисы
                var fileManager = provider.GetRequiredService<IFileManager>();
                var docLoader = provider.GetRequiredService<IDocumentLoader>();
                var exporter = provider.GetRequiredService<IExporter>();
                var annotator = provider.GetRequiredService<AnnotationGenerator>();

                // Пути к файлам
                string inputPath = @"D:\source\repos\DocumentComplianceChecker_HSEproject\add_files\input.docx";
                string outputPath = @"D:\source\repos\DocumentComplianceChecker_HSEproject\add_files\output.docx";
                string reportPath = @"D:\source\repos\DocumentComplianceChecker_HSEproject\add_files\report.txt";

                // Проверка существования файла
                if (!fileManager.FileExists(inputPath))
                {
                    LogTextBlock.Text = $"Файл {inputPath} не найден!";
                    return;
                }

                // Создаем копию документа для обработки
                using var doc = docLoader.CreateDocumentCopy(inputPath, outputPath);
                List<Error> errors;

                if (_useTemplate)
                {
                    var template = provider.GetRequiredService<Template>(); // Получаем шаблон из DI
                    var formatter = new FormattingTemplate(template); // Создаем объект для валидации с использованием шаблона
                    errors = new List<Error>();

                    // Обрабатываем каждый параграф
                    foreach (var paragraph in doc.MainDocumentPart.Document.Body.Descendants<Paragraph>())
                    {
                        var result = formatter.ValidateParagraph(paragraph);// Получаем ошибки для параграфа
                        errors.AddRange(result.Errors);// Добавляем ошибки в общий список
                    }
                }
                else
                {
                    // Если шаблон не выбран, используем стандартную валидацию
                    var validator = provider.GetRequiredService<IFormattingValidator>();
                    var formattingValidationResult = validator.Validate(doc);
                    errors = formattingValidationResult.Errors;// Просто присваиваем ошибки в переменную
                }

                // Преобразуем List<Error> в ValidationResult
                var validationResult = new MyValidationResult();
                validationResult.Errors.AddRange(errors);

                // Применяем аннотации к документу на основе ошибок
                annotator.ApplyAnnotations(doc, validationResult);
                // Экспортируем отчет
                exporter.ExportReport(validationResult, reportPath);

                LogTextBlock.Text = $"Проверка завершена. Результаты сохранены в:\n{outputPath}\n{reportPath}";
            }
            catch (Exception ex)
            {
                LogTextBlock.Text = $"Ошибка: {ex.Message}";
            }
        }
    }
}