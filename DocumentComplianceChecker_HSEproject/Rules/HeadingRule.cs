using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Services;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    public class HeadingRule : ValidationRule
    {
        private readonly FormattingTemplate _template;

        public HeadingRule(FormattingTemplate template)
        {
            _template = template;
        }

        public IEnumerable<Error> Validate(WordprocessingDocument document)
        {
            var errors = new List<Error>();
            var body = document.MainDocumentPart?.Document.Body;

            if (body == null)
                return errors;

            var headings = body.Descendants<Paragraph>()
                               .Where(p => p.ParagraphProperties?.ParagraphStyleId != null &&
                                           p.ParagraphProperties.ParagraphStyleId.Val.Value.StartsWith("Heading"));

            foreach (var heading in headings)
            {
                var styleId = heading.ParagraphProperties.ParagraphStyleId.Val.Value;
                var text = heading.InnerText;

                // Пример: проверка на отсутствие заглавных букв
                char? firstLetter = text.FirstOrDefault(char.IsLetter);

                if (firstLetter.HasValue && !char.IsUpper(firstLetter.Value))
                {
                    errors.Add(new HeadingError
                    {
                        Message = $"Заголовок '{text}' должен начинаться с заглавной буквы",
                        Paragraph = heading,
                        AnnotationType = AnnotationType.Warning
                    });
                }



                // Можно добавить другие проверки, например уровень стиля и отступ
            }

            return errors;
        }

        public override bool Validate(Paragraph paragraph, Run run = null)
        {
            throw new NotImplementedException();
        }
    }
}
