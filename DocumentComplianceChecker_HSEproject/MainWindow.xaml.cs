using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Rules;
using DocumentComplianceChecker_HSEproject.Services;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Win32;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.IO;
using System.Windows.Controls;

namespace DocumentComplianceChecker_HSEproject
{
    using MyValidationResult = DocumentComplianceChecker_HSEproject.Models.ValidationResult;

    public partial class MainWindow : Window
    {
        private readonly IServiceProvider provider;
        private bool useTemplate;
        private const string HelpFileName = "CHM_DocumentComplianceChecker.chm"; // Относительный путь

        public MainWindow()
        {
            InitializeComponent();

            // Регистрируем обработчик для клавиши F1
            this.KeyDown += MainWindow_KeyDown;
        }

        static void ConfigureServices(IServiceCollection services)
        {
            services.AddSingleton<IFileManager, FileManager>();
            services.AddTransient<IDocumentLoader, DocumentLoader>();
            services.AddTransient<IExporter, Exporter>();
            services.AddTransient<AnnotationGenerator>();

            // Регистрируем FormattingValidator с двумя списками правил
            services.AddScoped<IFormattingValidator>(provider =>
            {
                var paragraphRules = new List<IParagraphValidationRule>
                {
                    new NormalStyleRule(),
                    new Heading1Rule(),
                    new Heading2Rule(),
                    new Heading3Rule(),
                    new PageMarginRule()
                };
                var runRules = new List<IRunValidationRule>
                {
                    new NormalStyleRule(),
                    new Heading1Rule(),
                    new Heading2Rule(),
                    new Heading3Rule()
                };
                return new FormattingValidator(paragraphRules, runRules);
            });
        }

        private void CheckDocument_Click(object sender, RoutedEventArgs e)
        {
            // Настройка DI
            var services = new ServiceCollection();
            ConfigureServices(services); // Конфигурируем контейнер с необходимыми зависимостями

            try
            {
                // Определяем режим проверки на основе выбора пользователя
                useTemplate = TemplateRadio.IsChecked == true;
                //List<MyValidationRule> rules;

                if (useTemplate)
                {
                    var template = new Template(); // Создаём объект шаблона
                    services.AddSingleton(template); // Добавляем шаблон в DI контейнер
                }

                // Строим провайдер для получения зависимостей
                var provider = services.BuildServiceProvider();

                // Получаем сервисы
                var fileManager = provider.GetRequiredService<IFileManager>();
                var docLoader = provider.GetRequiredService<IDocumentLoader>();
                var exporter = provider.GetRequiredService<IExporter>();
                var annotator = provider.GetRequiredService<AnnotationGenerator>();

                // Получаем путь из текстового поля
                string inputPath = InputPathTextBox.Text.Trim();
                if (string.IsNullOrEmpty(inputPath))
                {
                    LogTextBlock.Text = "Пожалуйста, укажите путь к файлу";
                    return;
                }

                // Пути к файлам
                string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputFiles", "output.docx");
                string reportPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputFiles", "report.txt");
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

                // Проверка существования файла
                if (!fileManager.FileExists(inputPath))
                {
                    LogTextBlock.Text = $"Файл {inputPath} не найден";
                    return;
                }

                // Создаем копию документа для обработки
                using var doc = docLoader.CreateDocumentCopy(inputPath, outputPath);
                
                // Итоговый объект с результатами валидации
                MyValidationResult validationResult;
                List<string> errorMessages = new List<string>();

                HashSet<string> uniqueErrors = new HashSet<string>();

                if (useTemplate)
                {
                    var validator = provider.GetRequiredService<IFormattingValidator>();
                    validationResult = validator.Validate(doc);
                    //var template = provider.GetRequiredService<Template>(); // Получаем шаблон из DI
                    //var formatter = new FormattingTemplate(template); // Создаем объект для валидации с использованием шаблона

                    //validationResult = formatter.ValidateDocument(doc);

                    //// Проходим по всем ошибкам
                    //foreach (var error in validationResult.Errors)
                    //{
                    //    uniqueErrors.Add(error.Message); // Собираем уникальные ошибки
                    //}
                }
                else
                {
                    // Если шаблон не выбран, используем стандартную валидацию
                    var validator = provider.GetRequiredService<IFormattingValidator>();
                    validationResult = validator.Validate(doc);

                    // Проходим по всем ошибкам
                    foreach (var error in validationResult.Errors)
                    {
                        uniqueErrors.Add(error.Message); // Собираем уникальные ошибки
                    }
                }

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
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*",
                Title = "Выберите файл для проверки"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                InputPathTextBox.Text = openFileDialog.FileName;
            }
        }

        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Открываем CHM-файл
                Process.Start(new ProcessStartInfo(HelpFileName) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                // Показываем ошибку в логе
                LogTextBlock.Text = $"Ошибка при открытии справки: {ex.Message}";
            }
        }

        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F1)
            {
                try
                {
                    // Получаем текущий активный элемент
                    var focusedElement = Keyboard.FocusedElement as FrameworkElement;
                    string helpId = focusedElement?.Tag?.ToString();

                    if (!string.IsNullOrEmpty(helpId))
                    {
                        // Открываем CHM с указанием Help ID
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = "hh.exe",
                            Arguments = $"-mapid {helpId} {HelpFileName}",
                            UseShellExecute = true
                        });
                    }
                    else
                    {
                        // Если Help ID нет, открываем главную страницу
                        Process.Start(new ProcessStartInfo(HelpFileName) { UseShellExecute = true });
                    }
                    e.Handled = true; // Предотвращаем дальнейшую обработку F1
                }
                catch (Exception ex)
                {
                    LogTextBlock.Text = $"Ошибка при открытии контекстной справки: {ex.Message}";
                }
            }
        }
    }
}