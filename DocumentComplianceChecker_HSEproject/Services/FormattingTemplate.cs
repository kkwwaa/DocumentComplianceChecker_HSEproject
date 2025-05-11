using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Linq;

namespace DocumentComplianceChecker_HSEproject.Services
{
    internal class FormattingTemplate
    {
        private readonly Template _template;

        internal FormattingTemplate(Template template)
        {
            _template = template;
        }

        // Проверка одного абзаца
        public ValidationResult ValidateParagraph(Paragraph paragraph)
        {
            var result = new ValidationResult();

            foreach (var rule in _template.Rules)
            {
                bool passed = rule.Validate(paragraph);

                if (!passed)
                {
                    result.Errors.Add(new Error
                    {
                        RuleName = rule.GetType().Name,
                        Message = $"Нарушение правила: {rule.GetType().Name}"
                    });
                }
            }

            return result;
        }

        // Проверяет весь документ Word
        public ValidationResult ValidateDocument(WordprocessingDocument document)
        {
            var result = new ValidationResult();

            var paragraphs = document.MainDocumentPart?.Document?.Body?.Descendants<Paragraph>() ?? Enumerable.Empty<Paragraph>();

            foreach (var paragraph in paragraphs)
            {
                var paragraphResult = ValidateParagraph(paragraph);
                result.Errors.AddRange(paragraphResult.Errors);
            }

            return result;
        }
    }
}
