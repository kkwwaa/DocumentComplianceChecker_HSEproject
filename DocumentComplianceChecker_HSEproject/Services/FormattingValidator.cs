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
                var paragraph = paragraphs[i];
                var runs = paragraph.Elements<Run>().ToList();

                // Проверяем каждый Run в параграфе
                foreach (var run in runs)
                {
                    foreach (var rule in _rules)
                    {
                        if (!rule.Validate(paragraph, run)) // Теперь передаем конкретный Run
                        {
                            Console.WriteLine(GetRunText(run));

                            result.Errors.Add(new Error
                            {
                                ErrorType = rule.GetType().Name,
                                Message = rule.ErrorMessage,
                                ParagraphText = GetRunText(run),
                                ParagraphIndex = i,
                                TargetRun = run // Сохраняем ссылку на проблемный Run
                            });
                        }
                    }
                }
            }

            return result;
        }

        private string GetRunText(Run run)
        {
            return string.Concat(run.Elements<Text>().Select(t => t.Text));
        }
    }
}