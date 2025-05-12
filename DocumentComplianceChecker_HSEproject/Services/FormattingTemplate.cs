using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Services
{
    // Класс отвечает за применение набора правил форматирования (из шаблона) к отдельному абзацу
    internal class FormattingTemplate
    {
        // Приватное поле — шаблон, содержащий набор правил форматирования
        private readonly Template _template;

        // Конструктор, принимает объект шаблона и сохраняет его для дальнейшей валидации
        internal FormattingTemplate(Template template) // Template мог бы внутри себя строить список правил, которые он содержит
        {
            _template = template;
        }

        // Проверяет один абзац на соответствие всем правилам шаблона
        public ValidationResult ValidateParagraph(Paragraph paragraph) // сделать приватным
        {
            // Создаём объект результата валидации, который будет содержать ошибки
            var result = new ValidationResult();

            // Применяем каждое правило из шаблона к абзацу
            foreach (var rule in _template.Rules)
            {
                // Проверка: проходит ли абзац данное правило
                bool passed = rule.Validate(paragraph);

                // Если не проходит — добавляем ошибку в результат
                if (!passed)
                {
                    result.Errors.Add(new Error
                    {
                        RuleName = rule.GetType().Name, // Название класса правила
                        Message = $"Нарушение правила: {rule.GetType().Name}" // Сообщение об ошибке
                    });
                }
            }

            // Возвращаем итоговый результат с ошибками (если были)
            return result;
        }
    }
}
