using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentComplianceChecker_HSEproject.Models;

namespace DocumentComplianceChecker_HSEproject.Services
{
    public class AnnotationGenerator
    {
        public void ApplyAnnotations(WordprocessingDocument doc, ValidationResult validationResult)
        {
            if (validationResult.Errors.Count == 0) return;

            var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();

            foreach (var error in validationResult.Errors)
            {
                if (error.ParagraphIndex < paragraphs.Count)
                {
                    var paragraph = paragraphs[error.ParagraphIndex];

                    // 1. Подсветка конкретного текста с ошибкой
                    if (!string.IsNullOrEmpty(error.ParagraphText))
                    {
                        HighlightText(paragraph, error.ParagraphText);
                    }

                    // 2. Добавление комментария
                    AddComment(doc, paragraph, error);
                }
            }
        }

        private void HighlightText(Paragraph paragraph, string errorText)
        {
            // Ищем только Run'ы, которые полностью совпадают с текстом ошибки
            var exactMatchRuns = paragraph.Descendants<Run>()
                .Where(r => r.InnerText.Trim() == errorText.Trim())
                .ToList();

            foreach (var run in exactMatchRuns)
            {
                var runProperties = run.RunProperties ?? new RunProperties();
                runProperties.Highlight = new Highlight() { Val = HighlightColorValues.Red };
                run.RunProperties = runProperties;
            }
        }

        private void AddComment(WordprocessingDocument doc, Paragraph paragraph, Error error)
        {
            if (doc == null || doc.MainDocumentPart == null || paragraph == null || string.IsNullOrEmpty(error.Message))
            {
                throw new ArgumentException("Invalid arguments passed to AddComment");
            }

            // Получаем или создаем часть для комментариев
            var commentsPart = doc.MainDocumentPart.WordprocessingCommentsPart ??
                doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();

            commentsPart.Comments ??= new Comments();

            // Создаем комментарий с типом ошибки
            var comment = new Comment()
            {
                Id = GenerateUniqueCommentId(commentsPart),
                Author = "DocumentComplianceChecker",
            };

            comment.AppendChild(new Paragraph(
                new Run(
                    new Text($"{error.ErrorType}: {error.Message}")
                )));

            commentsPart.Comments.AppendChild(comment);

            // Привязываем комментарий к первому Run в параграфе
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun != null)
            {
                paragraph.InsertBefore(new CommentRangeStart { Id = comment.Id }, firstRun);
                paragraph.InsertAfter(new CommentRangeEnd { Id = comment.Id }, firstRun);
                paragraph.AppendChild(new Run(new CommentReference { Id = comment.Id }));
            }
        }

        private static string GenerateUniqueCommentId(WordprocessingCommentsPart commentsPart)
        {
            if (commentsPart?.Comments == null) return "1";

            return (commentsPart.Comments
                .Elements<Comment>()
                .Select(c => int.Parse(c.Id?.Value ?? "0"))
                .DefaultIfEmpty()
                .Max() + 1).ToString();
        }
    }
}