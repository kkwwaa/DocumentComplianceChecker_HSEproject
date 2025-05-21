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
            // Результат проверки, куда будут собираться ошибки
            var result = new ValidationResult();

            // Получаем все абзацы из основного тела документа
            var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();

            // Проходим по каждому абзацу
            for (int i = 0; i < paragraphs.Count; i++)
            {
                var paragraph = paragraphs[i]; // Текущий параграф

                // Проверка правил для параграфа
                foreach (var rule in _paragraphRules)
                {
                    if (!rule.ValidateParagraph(paragraph))
                    {
                        // Выводим текст первого Run для отладки (если есть)
                        var firstRun = paragraph.Elements<Run>().FirstOrDefault();
                        Console.WriteLine(GetRunText(firstRun));

                        // Добавляем информацию об ошибке в результат
                        result.Errors.Add(new Error
                        {
                            ErrorType = rule.GetType().Name,        // Тип правила, вызвавшего ошибку
                            Message = rule.ErrorMessage,           // Сообщение об ошибке
                            ParagraphText = GetRunText(firstRun),  // Текст первого Run
                            ParagraphIndex = i,                    // Индекс абзаца
                            TargetRun = firstRun                   // Ссылка на первый Run
                        });
                    }
                }

                // Извлекаем все Run из текущего абзаца
                var runs = paragraph.Elements<Run>().ToList();

                // Проверяем каждый Run в данном абзаце
                foreach (var run in runs)
                {
                    foreach (var rule in _runRules)
                    {
                        if (!rule.ValidateRun(paragraph, run))
                        {
                            // Выводим текст Run для отладки
                            Console.WriteLine(GetRunText(run));

                            // Добавляем информацию об ошибке в результат
                            result.Errors.Add(new Error
                            {
                                ErrorType = rule.GetType().Name,        // Тип правила, вызвавшего ошибку
                                Message = rule.ErrorMessage,           // Сообщение об ошибке
                                ParagraphText = GetRunText(run),       // Текст Run
                                ParagraphIndex = i,                    // Индекс абзаца
                                TargetRun = run                        // Ссылка на проблемный Run
                            });
                        }
                    }
                }
            }

            // Возвращаем итог со всеми найденными ошибками
            return result;
        }

        // Вспомогательный метод: извлекает объединённый текст из всех Text-элементов внутри Run
        private string GetRunText(Run run)
        {
            return run != null ? string.Concat(run.Elements<Text>().Select(t => t.Text)) : string.Empty;
        }
    }
}