using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

namespace DocumentComplianceChecker_HSEproject.Models
{
    public class Template
    {
        // Свойства шаблона
        public string Name { get; set; }  // Имя шаблона
        public string Font { get; set; }  // Шрифт, который используется в шаблоне
        public double FontSize { get; set; }  // Размер шрифта
        public string TemplatePath { get; set; }  // Путь к шаблону

        // Добавлены свойства из FormattingTemplate
        public bool BoldRequired { get; set; } = false;
        public bool ItalicRequired { get; set; } = false;
        public int LineSpacing { get; set; } = 2;
        public int HeadingFontSize { get; set; } = 16;
        public bool HeadingRule { get; internal set; }

        // Конструктор для создания шаблона
        public Template(string templatePath)
        {
            TemplatePath = templatePath;
            LoadTemplate(templatePath);  // Загружаем шаблон из файла
        }

        // Метод для загрузки шаблона и извлечения параметров
        private void LoadTemplate(string path)
        {
            try
            {
                // Открываем документ Word
                using (WordprocessingDocument doc = WordprocessingDocument.Open(path, false))
                {
                    // Извлекаем данные о шрифте и других параметрах из документа
                    var firstParagraph = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().FirstOrDefault();
                    if (firstParagraph != null)
                    {
                        var firstRun = firstParagraph.Elements<Run>().FirstOrDefault();
                        if (firstRun != null)
                        {
                            var runProperties = firstRun.Elements<RunProperties>().FirstOrDefault();
                            if (runProperties != null)
                            {
                                // Извлекаем шрифт
                                var font = runProperties.Elements<Font>().FirstOrDefault();
                                if (font != null)
                                {
                                    // пока не получилось реализовать логику
                                }

                                // Извлекаем размер шрифта
                                var fontSize = runProperties.Elements<FontSize>().FirstOrDefault();
                                if (fontSize != null)
                                {
                                    FontSize = double.Parse(fontSize.Val.Value) / 2; // Переводим размер шрифта
                                }

                                // Извлекаем другие свойства, такие как Bold или Italic
                                var bold = runProperties.Elements<Bold>().FirstOrDefault();
                                if (bold != null)
                                {
                                    BoldRequired = true;
                                }

                                var italic = runProperties.Elements<Italic>().FirstOrDefault();
                                if (italic != null)
                                {
                                    ItalicRequired = true;
                                }
                            }
                        }
                    }

                    // Извлекаем информацию о междустрочном интервале
                    var style = doc.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault();
                    if (style != null)
                    {
                        // Здесь можно извлечь дополнительные параметры, такие как междустрочный интервал
                        // Однако для этого нужен более сложный механизм, который анализирует стиль документа.
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке шаблона: {ex.Message}");
            }
        }

        // Метод для проверки валидности шаблона
        public bool IsValid()
        {
            // Простейшая проверка: шаблон должен иметь имя и шрифт
            return !string.IsNullOrEmpty(Name) && !string.IsNullOrEmpty(Font);
        }
    }
}
