using DocumentComplianceChecker_HSEproject.Interfaces;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Rules;
using DocumentComplianceChecker_HSEproject.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;

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
            ConfigureServices(services);

            try
            {
                var provider = services.BuildServiceProvider();

                var fileManager = provider.GetRequiredService<IFileManager>();
                var docLoader = provider.GetRequiredService<IDocumentLoader>();
                var exporter = provider.GetRequiredService<IExporter>();
                var annotator = provider.GetRequiredService<AnnotationGenerator>();

                string inputPath = InputPathTextBox.Text.Trim();
                if (string.IsNullOrEmpty(inputPath))
                {
                    LogTextBlock.Text = "Пожалуйста, укажите путь к файлу";
                    return;
                }

                string outputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputFiles");
                Directory.CreateDirectory(outputDir);

                string outputPath = Path.Combine(outputDir, "output.docx");
                string reportPath = Path.Combine(outputDir, "report.txt");

                if (!fileManager.FileExists(inputPath))
                {
                    LogTextBlock.Text = $"Файл {inputPath} не найден";
                    return;
                }

                using var doc = docLoader.CreateDocumentCopy(inputPath, outputPath);

                MyValidationResult validationResult;
                var uniqueErrors = new HashSet<string>();

                var validator = provider.GetRequiredService<IFormattingValidator>();
                validationResult = validator.Validate(doc);

                foreach (var error in validationResult.Errors)
                {
                    uniqueErrors.Add(error.Message);
                }

                annotator.ApplyAnnotations(doc, validationResult);
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
            var openFileDialog = new OpenFileDialog
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
            OpenHelpFile();
        }

        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F1)
            {
                OpenHelpFile();
                e.Handled = true;
            }
        }

        /// <summary>
        /// Пытается открыть CHM-файл справки. Если не найден – выводит ошибку в LogTextBlock.
        /// </summary>
        private void OpenHelpFile()
        {
            try
            {
                // Составляем полный путь до CHM, предполагая, что он лежит в директории с .exe
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string helpPath = Path.Combine(baseDir, HelpFileName);

                if (!File.Exists(helpPath))
                {
                    LogTextBlock.Text = $"Ошибка при открытии справки: файл '{HelpFileName}' не найден по пути '{helpPath}'.";
                    return;
                }

                // Открываем CHM-файл (основная страница)
                Process.Start(new ProcessStartInfo(helpPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                LogTextBlock.Text = $"Ошибка при открытии справки: {ex.Message}";
            }
        }
    }
}