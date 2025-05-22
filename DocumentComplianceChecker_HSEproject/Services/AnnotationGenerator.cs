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

            // Группировка ошибок по параграфам
            var errorsByParagraph = validationResult.Errors
                .Where(e => e.ParagraphIndex >= 0 && e.ParagraphIndex < paragraphs.Count)
                .GroupBy(e => e.ParagraphIndex);

            foreach (var paragraphGroup in errorsByParagraph)
            {
                int paragraphIndex = paragraphGroup.Key;
                var paragraph = paragraphs[paragraphIndex];
                var groupedErrors = paragraphGroup.ToList();

                var commentText = $"Параграф {paragraphIndex + 1}:\n" + string.Join("\n", groupedErrors
                    .Select(e => $"- [{e.ErrorType}] {e.Message}")
                    .Distinct());

                AddComment(doc, paragraph, commentText);
            }
        }

        private void AddComment(WordprocessingDocument doc, Paragraph paragraph, string commentText)
        {
            if (doc == null || doc.MainDocumentPart == null || paragraph == null || string.IsNullOrEmpty(commentText))
                return;

            var commentsPart = doc.MainDocumentPart.WordprocessingCommentsPart
                ?? doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments ??= new Comments();

            string commentId = GenerateUniqueCommentId(commentsPart);

            var comment = new Comment
            {
                Id = commentId,
                Author = "DocumentComplianceChecker",
                Date = DateTime.Now
            };
            comment.AppendChild(new Paragraph(new Run(new Text(commentText))));
            commentsPart.Comments.AppendChild(comment);

            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun == null)
            {
                firstRun = new Run(new Text(""));
                paragraph.AppendChild(firstRun);
            }

            paragraph.InsertBefore(new CommentRangeStart { Id = commentId }, firstRun);
            paragraph.InsertAfter(new CommentRangeEnd { Id = commentId }, firstRun);
            paragraph.InsertAfter(new Run(new CommentReference { Id = commentId }), firstRun);

            commentsPart.Comments.Save();
        }

        private static string GenerateUniqueCommentId(WordprocessingCommentsPart commentsPart)
        {
            var maxId = commentsPart.Comments.Elements<Comment>()
                .Select(c => int.TryParse(c.Id?.Value, out int val) ? val : 0)
                .DefaultIfEmpty(0)
                .Max();

            return (maxId + 1).ToString();
        }
    }
}