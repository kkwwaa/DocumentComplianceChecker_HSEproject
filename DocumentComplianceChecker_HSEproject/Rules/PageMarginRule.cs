using DocumentComplianceChecker_HSEproject.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    public class PageMarginRule : ValidationRule
    {
        // Допустимые значения полей в сантиметрах
        private const double TopCm = 2.0;
        private const double BottomCm = 2.0;
        private const double LeftCm = 3.0;
        private const double RightCm = 1.5;

        // Преобразование см в twentieths of a point (единицы OpenXML)
        private static int CmToTwips(double cm) => (int)(cm * 567); // 1 см ≈ 567 twips

        public override bool Validate(Paragraph paragraph, Run run = null)
        {
            // Получаем доступ к документу через родительскую иерархию
            var document = paragraph.Ancestors<Body>()
                .FirstOrDefault()?.Parent as Document;
            if (document == null) return true;

            var sectionProps = paragraph.Descendants<SectionProperties>().FirstOrDefault()
                ?? document.Body?.Elements<SectionProperties>().FirstOrDefault();

            var pageMargin = sectionProps?.GetFirstChild<PageMargin>();
            if (pageMargin == null) return false;

            // Сравнение с допустимыми значениями
            return pageMargin.Top == CmToTwips(TopCm)
                && pageMargin.Bottom == CmToTwips(BottomCm)
                && pageMargin.Left == CmToTwips(LeftCm)
                && pageMargin.Right == CmToTwips(RightCm);
        }
    }
}