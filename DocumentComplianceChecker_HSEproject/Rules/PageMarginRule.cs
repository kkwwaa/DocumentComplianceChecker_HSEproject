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

        // Обратная конвертация: твипы -> сантиметры (для вывода)
        private static double TwipsToCm(int twips) => twips / 567.0;

        // Проверка полей страницы
        public bool ValidateParagraph(Paragraph paragraph)
        {
            // Получаем доступ к документу через параграф
            var document = paragraph.Ancestors<Body>()?.FirstOrDefault()?.Parent as Document;
            if (document == null)
            {
                Console.WriteLine("[PageMarginRule] Не удалось получить объект Document из параграфа.");
                return false;
            }

            // Ищем свойства раздела
            var sectionProps = paragraph
                .Descendants<SectionProperties>()
                .FirstOrDefault()
                ?? document.Body?.Elements<SectionProperties>().FirstOrDefault();

            if (sectionProps == null)
            {
                Console.WriteLine("[PageMarginRule] SectionProperties не найдены.");
                return false;
            }

            // Получаем объект PageMargin
            var margin = sectionProps.GetFirstChild<PageMargin>();
            if (margin == null)
            {
                Console.WriteLine("[PageMarginRule] PageMargin не задан.");
                return false;
            }

            // Текущие значения полей в твипсах
            int topTwips = margin.Top?.Value ?? 0;
            int bottomTwips = margin.Bottom?.Value ?? 0;
            int leftTwips = Convert.ToInt32(margin.Left?.Value ?? 0u);
            int rightTwips = Convert.ToInt32(margin.Right?.Value ?? 0u);

            // Конвертируем в сантиметры
            double topCm = TwipsToCm(topTwips);
            double bottomCm = TwipsToCm(bottomTwips);
            double leftCm = TwipsToCm(leftTwips);
            double rightCm = TwipsToCm(rightTwips);

            // Выводим значения полей
            Console.WriteLine("[PageMarginRule] Текущие поля страницы:");
            Console.WriteLine($"  Top:    {topTwips} твипс ({topCm:F2} см)");
            Console.WriteLine($"  Bottom: {bottomTwips} твипс ({bottomCm:F2} см)");
            Console.WriteLine($"  Left:   {leftTwips} твипс ({leftCm:F2} см)");
            Console.WriteLine($"  Right:  {rightTwips} твипс ({rightCm:F2} см)");

            // Параметры, которые мы ожидаем:
            // Top = 2 см  => CmToTwips(2.0)  = 1134 твипса
            // Bottom = 2 см => 1134 твипса
            // Left = 3 см => 1701 твипс
            // Right = 1.5 см => 850 твипсов (≈850.5, но Int приближает до 850)
            int expectedTopTwips = CmToTwips(2.0);
            int expectedBottomTwips = CmToTwips(2.0);
            int expectedLeftTwips = CmToTwips(3.0);
            int expectedRightTwips = CmToTwips(1.5);

            // Проверка корректности
            bool isTopOk = topTwips == expectedTopTwips;
            bool isBottomOk = bottomTwips == expectedBottomTwips;
            bool isLeftOk = leftTwips == expectedLeftTwips;
            bool isRightOk = rightTwips == expectedRightTwips;

            if (!isTopOk || !isBottomOk || !isLeftOk || !isRightOk)
            {
                Console.WriteLine("[PageMarginRule] Обнаружены несоответствия полей страницы.");
                Console.WriteLine($"  Ожидаемые значения (twips): Top={expectedTopTwips}, Bottom={expectedBottomTwips}, Left={expectedLeftTwips}, Right={expectedRightTwips}");
                return false;
            }

            Console.WriteLine("[PageMarginRule] Поля страницы соответствуют требованиям.");
            return true;
        }
    }
}