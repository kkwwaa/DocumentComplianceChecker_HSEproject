using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Interfaces;

namespace DocumentComplianceChecker_HSEproject.Services
{
    public class FormattingValidator : IFormattingValidator
    {
        public List<Error> Validate(WordprocessingDocument doc)
        {
            List<Error> errors = new List<Error>();
            var body = doc.MainDocumentPart?.Document.Body;

            if (body == null) return errors;

            // Получаем стиль документа по умолчанию
            var defaultFont = doc.MainDocumentPart.StyleDefinitionsPart?
                               .Styles.Elements<Style>()
                               .FirstOrDefault(s => s.Type == StyleValues.Paragraph && s.Default)?
                               .StyleRunProperties?.RunFonts?.Ascii?.Value ?? "Times New Roman";

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    var runProperties = run.RunProperties;

                    // Проверяем шрифт несколькими способами
                    var fontName = runProperties?.RunFonts?.Ascii?.Value
                                  ?? runProperties?.RunFonts?.HighAnsi?.Value
                                  ?? runProperties?.RunFonts?.ComplexScript?.Value
                                  ?? defaultFont;

                    if (fontName != "Times New Roman")
                    {
                        errors.Add(new Error
                        {
                            ErrorType = "InvalidFont",
                            Message = $"Недопустимый шрифт: '{fontName}'. Должен быть 'Times New Roman'.",
                            ParagraphText = paragraph.InnerText
                        });
                    }
                }
            }
            return errors;
        }
    }
}