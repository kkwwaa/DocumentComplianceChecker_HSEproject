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
            List<int> pageStartParagraphIndices = new() { 0 }; // первая страница всегда начинается с первого абзаца

            for (int i = 0; i < paragraphs.Count; i++)
            {
                var para = paragraphs[i];

                // 1. Вставлен ручной разрыв страницы: <w:br w:type="page"/>
                bool hasPageBreak = para.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page);

                // 2. Свойство абзаца: <w:pageBreakBefore/>
                bool hasPageBreakBefore = para.ParagraphProperties?.PageBreakBefore != null;

                // 3. Раздел начинается с новой страницы: <w:sectPr><w:type w:val="nextPage"/>
                bool hasSectionBreakWithNewPage = para.Descendants<SectionProperties>()
                    .Any(sp => sp.GetFirstChild<SectionType>()?.Val?.Value == SectionMarkValues.NextPage);

                if (hasPageBreak || hasPageBreakBefore || hasSectionBreakWithNewPage)
                {
                    // следующая страница начнётся со следующего абзаца
                    if (i + 1 < paragraphs.Count)
                        pageStartParagraphIndices.Add(i + 1);
                }
            }

            // Функция определения номера страницы по индексу абзаца
            int GetPageNumber(int paragraphIndex)
            {
                for (int i = pageStartParagraphIndices.Count - 1; i >= 0; i--)
                {
                    if (paragraphIndex >= pageStartParagraphIndices[i])
                        return i + 1; // страницы нумеруются с 1
                }
                return 1;
            }

            // Группировка ошибок по номеру страницы
            var errorsByPage = validationResult.Errors
                .Where(e => e.ParagraphIndex >= 0 && e.ParagraphIndex < paragraphs.Count)
                .GroupBy(e => GetPageNumber(e.ParagraphIndex));

            foreach (var pageGroup in errorsByPage)
            {
                int pageIndex = pageGroup.Key;
                var groupedErrors = pageGroup.ToList();

                var representativeParagraphIndex = groupedErrors
                    .Select(e => e.ParagraphIndex)
                    .Where(i => i >= 0 && i < paragraphs.Count)
                    .Min();

                var paragraph = paragraphs[representativeParagraphIndex];

                var commentText =      string.Join("\n", groupedErrors
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

            var firstRun = paragraph.Elements<Run>().FirstOrDefault() ?? paragraph.AppendChild(new Run());
            paragraph.InsertBefore(new CommentRangeStart { Id = commentId }, firstRun);
            paragraph.InsertAfter(new CommentRangeEnd { Id = commentId }, firstRun);
            paragraph.AppendChild(new Run(new CommentReference { Id = commentId }));
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