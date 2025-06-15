using NUnit.Framework;
using TechTalk.SpecFlow;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentComplianceChecker_HSEproject.Models;
using DocumentComplianceChecker_HSEproject.Rules;
using DocumentComplianceChecker_HSEproject.Services;
using System.IO;
using DocumentFormat.OpenXml;
using static BasicRules;

namespace DocumentComplianceChecker_HSEproject.Specs.Steps
{
    [Binding]
    public class ValidationSteps
    {
        private ValidationResult _validationResult;
        private WordprocessingDocument _document;

        [Given(@"Я создаю документ с заголовком 1 уровня с некорректным форматированием")]
        public void GivenHeading1Incorrect()
        {
            _document = CreateDocument("2", 200, 300, "200", 16, "Times New Roman", true);
        }

        [Given(@"Я создаю документ с заголовком 2 уровня с неверными интервалами")]
        public void GivenHeading2Incorrect()
        {
            _document = CreateDocument("3", 100, 100, "100", 14, "Times New Roman");
        }

        [Given(@"Я создаю документ с заголовком 3 уровня с неверным шрифтом")]
        public void GivenHeading3IncorrectFont()
        {
            _document = CreateDocument("4", 120, 40, null, 13, "Arial", false, true);
        }

        [Given(@"Я создаю документ с обычным абзацем с неправильным выравниванием")]
        public void GivenNormalWithWrongJustification()
        {
            var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var paragraph = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = "Normal" },
                        new Justification { Val = JustificationValues.Center },
                        new SpacingBetweenLines { Line = "400" },
                        new Indentation { FirstLine = "709" }
                    ),
                    new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "Times New Roman" },
                            new FontSize { Val = "26" }
                        ),
                        new Text("Обычный текст с неправильным выравниванием")
                    )
                );

                mainPart.Document.Body.Append(paragraph);
                mainPart.Document.Save();
            }

            _document = WordprocessingDocument.Open(stream, false);
        }

        [When(@"Я запускаю валидацию документа с правилами (.*)")]
        public void WhenValidateWithRules(string ruleNames)
        {
            var rules = ruleNames.Split(',').Select(r => r.Trim()).ToList();

            var paragraphRules = new List<Interfaces.IParagraphValidationRule>();
            var runRules = new List<Interfaces.IRunValidationRule>();

            foreach (var rule in rules)
            {
                switch (rule)
                {
                    case "Heading1Rule":
                        paragraphRules.Add(new Heading1Rule());
                        runRules.Add(new Heading1Rule());
                        break;
                    case "Heading2Rule":
                        paragraphRules.Add(new Heading2Rule());
                        runRules.Add(new Heading2Rule());
                        break;
                    case "Heading3Rule":
                        paragraphRules.Add(new Heading3Rule());
                        runRules.Add(new Heading3Rule());
                        break;
                    case "NormalStyleRule":
                        paragraphRules.Add(new NormalStyleRule());
                        runRules.Add(new NormalStyleRule());
                        break;
                }
            }

            var validator = new FormattingValidator(paragraphRules, runRules);
            _validationResult = validator.Validate(_document);
        }

        [Then(@"В списке ошибок должна быть ошибка с типом \'(.*)'\")]
        public void ThenErrorShouldContain(string expectedRule)
        {
            var hasError = _validationResult.Errors.Any(e => e.ErrorType.Contains(expectedRule));
            Assert.IsTrue(hasError, $"Ожидалась ошибка от правила {expectedRule}, но она не найдена.");
        }

        private WordprocessingDocument CreateDocument(string styleId, int before, int after, string indent, int sizePt, string font, bool pageBreak = false, bool bold = false)
        {
            var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var props = new ParagraphProperties(
                    new ParagraphStyleId { Val = styleId },
                    new SpacingBetweenLines { Before = before.ToString(), After = after.ToString() }
                );

                if (!string.IsNullOrEmpty(indent))
                    props.Indentation = new Indentation { FirstLine = indent };

                if (pageBreak)
                    props.Append(new PageBreakBefore());

                var runProps = new RunProperties(
                    new RunFonts { Ascii = font },
                    new FontSize { Val = (sizePt * 2).ToString() }
                );

                if (bold)
                    runProps.Bold = new Bold();

                var paragraph = new Paragraph(
                    props,
                    new Run(runProps, new Text("Test paragraph"))
                );

                mainPart.Document.Body.Append(paragraph);
                mainPart.Document.Save();
            }

            return WordprocessingDocument.Open(stream, false);
        }
    }
}