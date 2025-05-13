using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Interfaces;

namespace DocumentComplianceChecker_HSEproject.Services
{
    // Класс, реализующий интерфейс IFormattingValidator, отвечает за проверку форматирования документа
    public class FormattingValidator : IFormattingValidator
    {
        // Список правил форматирования, которые будут применяться к каждому элементу документа
        private readonly List<IValidationRule> _rules;

        // Конструктор: принимает список правил и сохраняет для использования валидации
        public FormattingValidator(List<IValidationRule> rules)
        {
            _rules = rules;
        }

        // Метод валидации: обходит все абзацы документа и применяет к каждому Run все правила
        public ValidationResult Validate(WordprocessingDocument doc)
        {
            // Результат проверки, куда будут собираться ошибки
            var result = new ValidationResult();

            // Получаем все абзацы из основного тела документа
            var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();

            // Проходим по каждому абзацу
            for (int i = 0; i < paragraphs.Count; i++)
            {
                var paragraph = paragraphs[i];

                // Извлекаем все Run'ы (фрагменты текста с оформлением) из текущего абзаца
                var runs = paragraph.Elements<Run>().ToList();

                // Проверяем каждый Run в данном абзаце
                foreach (var run in runs)
                {
                    // Применяем каждое правило к текущему Run (и параграфу, если необходимо)
                    foreach (var rule in _rules)
                    {
                        // Если правило не выполнено
                        if (!rule.RuleValidator(paragraph, run))
                        {
                            // Выводим текст Run'а в консоль для отладки
                            Console.WriteLine(GetRunText(run));

                            // Добавляем информацию об ошибке в результат
                            result.Errors.Add(new Error
                            {
                                ErrorType = rule.GetType().Name,        // Тип правила, вызвавшего ошибку
                                Message = rule.ErrorMessage,           // Сообщение об ошибке
                                ParagraphText = GetRunText(run),       // Текст, в котором обнаружена ошибка
                                ParagraphIndex = i,                    // Индекс абзаца, в котором ошибка
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
            return string.Concat(run.Elements<Text>().Select(t => t.Text));
        }
    }
}