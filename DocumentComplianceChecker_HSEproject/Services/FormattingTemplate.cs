namespace DocumentComplianceChecker_HSEproject.Services
{
    public class FormattingTemplate
    {
        public string Name { get; set; } = "Default Template";
        public string FontName { get; set; } = "Times New Roman";
        public double FontSize { get; set; } = 14;
        public bool BoldRequired { get; set; } = false;
        public bool ItalicRequired { get; set; } = false;
        public int LineSpacing { get; set; } = 2;
        public int HeadingFontSize { get; set; } = 16;
        public bool HeadingRule { get; internal set; }
    }
}