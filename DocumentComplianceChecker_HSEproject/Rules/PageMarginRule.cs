using DocumentFormat.OpenXml.Wordprocessing;
using DocumentComplianceChecker_HSEproject.Interfaces;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    // Класс для проверки полей страницы, реализует только IParagraphValidationRule
    internal class PageMarginRule : IParagraphValidationRule
    {
        // Сообщение об ошибке, используется в ValidationResult
        public string ErrorMessage => "Неверные поля страницы.";

        // Конвертация сантиметров в твипсы для проверки
        private static int CmToTwips(double cm) => (int)(cm * 567);

        // Проверка полей страницы
        public bool ValidateParagraph(Paragraph paragraph)
        {
            // Получаем доступ к документу через параграф
            var document = paragraph.Ancestors<Body>()?.FirstOrDefault()?.Parent as Document;
            if (document == null) return false;

            // Ищем свойства раздела
            var sectionProps = paragraph.Descendants<SectionProperties>().FirstOrDefault()
                ?? document.Body?.Elements<SectionProperties>().FirstOrDefault();

            // Проверяем поля страницы
            var margin = sectionProps?.GetFirstChild<PageMargin>();
            if (margin == null) return false;

            if (margin.Top != CmToTwips(2.0)) return false;
            if (margin.Bottom != CmToTwips(2.0)) return false;
            if (margin.Left != CmToTwips(3.0)) return false;
            if (margin.Right != CmToTwips(1.5)) return false;
            if (margin.Header != CmToTwips(1.5)) return false;
            if (margin.Footer != CmToTwips(1.25)) return false;

            return true;
        }
    }
}