using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentComplianceChecker_HSEproject.Models;
using System.Globalization;

namespace DocumentComplianceChecker_HSEproject.Services
{
    public class AnnotationGenerator
    {
        public void ApplyAnnotations(WordprocessingDocument doc, List<Error> errors)
        {
            if (errors.Count == 0) return; // Если ошибок нет, выходим

            var body = doc.MainDocumentPart.Document.Body; // Получаем тело документа

            foreach (var error in errors) // Перебираем все ошибки
            {
                // Ищем параграф, содержащий текст с ошибкой
                // (необходима другая модель ошибки, чтобы искать не по тексту, а более точно)
                var targetParagraph = body
                    .Descendants<Paragraph>()
                    .FirstOrDefault(p => p.InnerText.Contains(error.ParagraphText));

                if (targetParagraph != null) // Если нашли параграф
                {
                    HighlightText(targetParagraph, error.ParagraphText); // Подсвечиваем текст
                    AddComment(doc, targetParagraph, error.Message); // Добавляем комментарий
                }
            }
        }

        private void HighlightText(Paragraph paragraph, string errorText)
        {
            // Находим все Run элементы в параграфе, содержащие текст с ошибкой
            var runsWithError = paragraph.Descendants<Run>()
                                        .Where(r => r.InnerText.Contains(errorText))
                                        .ToList();

            foreach (var run in runsWithError)
            {
                // Получаем или создаем свойства форматирования (RunProperties)
                var runProperties = run.RunProperties ?? new RunProperties();

                // Добавляем красное выделение
                runProperties.Highlight = new Highlight() { Val = HighlightColorValues.Red };

                // Обновляем свойства Run
                run.RunProperties = runProperties;
            }
        }

        private void AddComment(WordprocessingDocument doc, Paragraph paragraph, string commentText)
        {
            if (doc == null || doc.MainDocumentPart == null || paragraph == null || string.IsNullOrEmpty(commentText))
            {
                throw new ArgumentException("Invalid arguments passed to AddComment");
            }

            // Получаем или создаем часть для комментариев
            var commentsPart = doc.MainDocumentPart.WordprocessingCommentsPart;
            if (commentsPart == null)
            {
                commentsPart = doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                commentsPart.Comments = new Comments(); // Инициализируем коллекцию комментариев
            }

            // Убеждаемся, что коллекция Comments существует
            if (commentsPart.Comments == null)
            {
                commentsPart.Comments = new Comments();
            }

            // Создаем комментарий
            var comment = new Comment()
            {
                Id = GenerateUniqueCommentId(commentsPart), // Генерация уникального ID
                Author = "DocumentComplianceChecker",
            };

            // Создаем содержимое комментария
            var commentParagraph = new Paragraph(
                new Run(
                    new Text(commentText)
                )
            );

            // Добавляем параграф в комментарий
            comment.AppendChild(commentParagraph);

            // Добавляем комментарий в документ
            commentsPart.Comments.AppendChild(comment);

            // Размечаем место в тексте
            paragraph.InsertBefore(new CommentRangeStart { Id = comment.Id }, paragraph.GetFirstChild<Run>());
            paragraph.InsertAfter(new CommentRangeEnd { Id = comment.Id }, paragraph.Elements<Run>().Last());
            paragraph.AppendChild(new CommentReference { Id = comment.Id });
        }

        // Генерация уникального ID для комментария
        private static string GenerateUniqueCommentId(WordprocessingCommentsPart commentsPart)
        {
            if (commentsPart?.Comments == null)
            {
                return "1"; // Возвращаем значение по умолчанию, если нет комментариев
            }

            int maxId = 0;
            foreach (var existingComment in commentsPart.Comments.Elements<Comment>())
            {
                if (int.TryParse(existingComment.Id?.Value, out int id) && id > maxId)
                {
                    maxId = id;
                }
            }
            return (maxId + 1).ToString();
        }
    }
}
