using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Models
{
    public class TemplateManager
    {
        private List<Template> templates = new List<Template>();

        // Метод для загрузки шаблона по пути
        public Template LoadTemplate(string path)
        {
            if (File.Exists(path))
            {
                Template template = new Template(path);
                if (IsTemplateValid(template))
                {
                    templates.Add(template);
                    return template;
                }
                else
                {
                    Console.WriteLine("Шаблон невалиден.");
                    return null;
                }
            }
            else
            {
                Console.WriteLine($"Файл по пути {path} не существует.");
                return null;
            }
        }

        // Метод для проверки валидности шаблона
        public bool IsTemplateValid(Template template)
        {
            return template != null && !string.IsNullOrEmpty(template.Name) && !string.IsNullOrEmpty(template.Font);
        }

        // Метод для выбора шаблона из списка
        public Template SelectTemplate(string templateName)
        {
            return templates.FirstOrDefault(t => t.Name.Equals(templateName, StringComparison.OrdinalIgnoreCase));
        }

        // Метод для сравнения документа с шаблоном
        public List<string> CompareToTemplate(WordprocessingDocument doc, Template template)
        {
            List<string> differences = new List<string>();

            try
            {
                // Перебираем все параграфы документа и сравниваем их шрифт и размер с шаблоном
                var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>();

                foreach (var paragraph in paragraphs)
                {
                    var runs = paragraph.Elements<DocumentFormat.OpenXml.Spreadsheet.Run>();

                    foreach (var run in runs)
                    {
                        var runProperties = run.Elements<DocumentFormat.OpenXml.Spreadsheet.RunProperties>().FirstOrDefault();
                        if (runProperties != null)
                        {
                            // Извлекаем шрифт из документа
                            var fontNameElement = runProperties.Elements<FontName>().FirstOrDefault();
                            if (fontNameElement != null)
                            {
                                string fontName = fontNameElement.Val?.Value;
                                if (fontName != template.Font)
                                {
                                    differences.Add($"Шрифт отличается от шаблона: {fontName}");
                                }
                            }

                            // Извлекаем размер шрифта
                            var fontSizeElement = runProperties.Elements<DocumentFormat.OpenXml.Wordprocessing.FontSize>().FirstOrDefault();
                            if (fontSizeElement != null)
                            {
                                double fontSize = double.Parse(fontSizeElement.Val.Value) / 2;  // Переводим в стандартный размер
                                if (fontSize != template.FontSize)
                                {
                                    differences.Add($"Размер шрифта отличается от шаблона: {fontSize}");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при сравнении с шаблоном: {ex.Message}");
            }

            return differences;
        }
    }
}
