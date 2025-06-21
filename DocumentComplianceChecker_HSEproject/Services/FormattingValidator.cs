using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Interfaces;

namespace DocumentComplianceChecker_HSEproject.Services
{
    // Класс, реализующий интерфейс IFormattingValidator, отвечает за проверку форматирования документа
    public class FormattingValidator : IFormattingValidator
    {
        // Список правил форматирования, разделённых на два типа
        private readonly List<IParagraphValidationRule> _paragraphRules;
        private readonly List<IRunValidationRule> _runRules;

        // Конструктор: принимает списки правил и сохраняет для использования валидации
        public FormattingValidator(List<IParagraphValidationRule> paragraphRules, List<IRunValidationRule> runRules)
        {
            _paragraphRules = paragraphRules ?? new List<IParagraphValidationRule>();
            _runRules = runRules ?? new List<IRunValidationRule>();
        }

        // Метод валидации: обходит все абзацы документа и применяет правила
        public ValidationResult Validate(WordprocessingDocument doc)
        {
            var result = new ValidationResult();
            if (doc?.MainDocumentPart?.Document?.Body == null)
            {
                Console.WriteLine("Validate: Document body is null or invalid.");
                return result;
            }

            var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();
            Console.WriteLine($"Validate: Found {paragraphs.Count} paragraphs.");

            for (int i = 0; i < paragraphs.Count; i++)
            {
                var paragraph = paragraphs[i];
                Console.WriteLine($"Validate: Processing paragraph {i}");

                foreach (var rule in _paragraphRules)
                {
                    Console.WriteLine($"Validate: Applying paragraph rule {rule.GetType().Name}");
                    if (!rule.ValidateParagraph(paragraph))
                    {
                        var firstRun = paragraph.Elements<Run>().FirstOrDefault();
                        if (firstRun != null)
                        {
                            var runText = GetRunText(firstRun);
                            Console.WriteLine($"Validate: First Run text: {runText}");
                            result.Errors.Add(new Error
                            {
                                ErrorType = rule.GetType().Name,
                                Message = rule.ErrorMessage,
                                ParagraphText = runText,
                                ParagraphIndex = i,
                                TargetRun = firstRun
                            });
                        }
                        else
                        {
                            Console.WriteLine($"Validate: Paragraph {i} has no Runs, skipping annotation.");
                            result.Errors.Add(new Error
                            {
                                ErrorType = rule.GetType().Name,
                                Message = rule.ErrorMessage,
                                ParagraphText = "No Runs in paragraph",
                                ParagraphIndex = i,
                                TargetRun = null
                            });
                        }
                    }
                }

                var runs = paragraph.Elements<Run>().ToList();
                Console.WriteLine($"Validate: Found {runs.Count} Runs in paragraph {i}");
                foreach (var run in runs)
                {
                    foreach (var rule in _runRules)
                    {
                        Console.WriteLine($"Validate: Applying Run rule {rule.GetType().Name} to Run");
                        if (!rule.ValidateRun(paragraph, run))
                        {
                            var runText = GetRunText(run);
                            Console.WriteLine($"Validate: Run text: {runText}");
                            result.Errors.Add(new Error
                            {
                                ErrorType = rule.GetType().Name,
                                Message = rule.ErrorMessage,
                                ParagraphText = runText,
                                ParagraphIndex = i,
                                TargetRun = run
                            });
                        }
                    }
                }
            }
            Console.WriteLine("Validate: Validation completed.");
            return result;
        }

        // Вспомогательный метод: извлекает объединённый текст из всех Text-элементов внутри Run
        private string GetRunText(Run run)
        {
            return run != null ? string.Concat(run.Elements<Text>().Select(t => t.Text)) : string.Empty;
        }
    }
}