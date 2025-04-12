using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Interfaces;

namespace DocumentComplianceChecker_HSEproject.Services
{
    public class FormattingValidator : IFormattingValidator
    {
        private readonly List<ValidationRule> _rules;

        public FormattingValidator(List<ValidationRule> rules)
        {
            _rules = rules;
        }

        public ValidationResult Validate(WordprocessingDocument doc)
        {
            var result = new ValidationResult();
            var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();

            for (int i = 0; i < paragraphs.Count; i++)
            {
                foreach (var rule in _rules)
                {
                    if (!rule.Validate(paragraphs[i]))
                    {
                        result.Errors.Add(new Error
                        {
                            ErrorType = rule.GetType().Name,
                            Message = rule.ErrorMessage,
                            ParagraphText = paragraphs[i].InnerText,
                            ParagraphIndex = i
                        });
                    }
                }
            }

            return result;
        }
    }
}