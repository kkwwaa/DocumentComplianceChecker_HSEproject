using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace DocumentComplianceChecker_HSEproject.Rules
{
    internal class NormalStyleRule : StyleBasedRule
    {
        private readonly StringBuilder _errorMessages = new StringBuilder();

        public override string ErrorMessage
        {
            get
            {
                return _errorMessages.Length > 0
                    ? $"Нарушения в обычном тексте:\n{_errorMessages}"
                    : "Нарушения в обычном тексте.";
            }
        }

        public override bool ValidateParagraph(Paragraph paragraph)
        {
            _errorMessages.Clear();

            // Проверка параграфа
            if (paragraph == null)
            {
                _errorMessages.AppendLine("- Параграф отсутствует.");
                return false;
            }

            // Получаем стиль параграфа
            var styleId = GetStyleId(paragraph);
            Console.WriteLine($"[NormalStyleRule] Стиль параграфа: {styleId}");

            if (styleId != "Normal") return true;

            // Получаем свойства параграфа
            var props = paragraph.ParagraphProperties;
            if (props == null)
            {
                _errorMessages.AppendLine("- Свойства параграфа отсутствуют.");
            }

            // Проверка межстрочного интервала (1.5 строки = 18 pt)
            var spacing = props?.SpacingBetweenLines;
            if (spacing == null || spacing.Line == null)
            {
                _errorMessages.AppendLine("- Межстрочный интервал не задан.");
            }
            else
            {
                if (!int.TryParse(spacing.Line.Value, out var lineVal))
                {
                    _errorMessages.AppendLine("- Не удалось определить межстрочный интервал.");
                }
                else if (Math.Abs(lineVal / 20.0 - 18.0) > 1.0)
                {
                    _errorMessages.AppendLine($"- Межстрочный интервал {lineVal / 20.0:F1} pt, ожидается 18.0 pt.");
                }
            }

            // Проверка красной строки (1.25 см)
            var indent = props?.Indentation?.FirstLine;
            if (string.IsNullOrEmpty(indent))
            {
                _errorMessages.AppendLine("- Красная строка не задана.");
            }
            else
            {
                if (!double.TryParse(indent, out var indentTwips))
                {
                    _errorMessages.AppendLine("- Не удалось определить величину красной строки.");
                }
                else if (Math.Abs(indentTwips / 567.0 - 1.25) > 0.1)
                {
                    _errorMessages.AppendLine($"- Красная строка {indentTwips / 567.0:F2} см, ожидается 1.25 см.");
                }
            }

            // Проверка выравнивания (по ширине)
            var justification = props?.Justification?.Val;
            if (justification == null)
            {
                _errorMessages.AppendLine("- Выравнивание не задано.");
            }
            else if (justification != JustificationValues.Both)
            {
                _errorMessages.AppendLine($"- Выравнивание: {justification}, ожидается по ширине.");
            }

            // Проверка отступов (должны быть 0)
            var indentation = props?.Indentation;
            if (indentation?.Left != null && indentation.Left.Value != "0")
            {
                _errorMessages.AppendLine("- Левый отступ должен быть 0.");
            }
            if (indentation?.Right != null && indentation.Right.Value != "0")
            {
                _errorMessages.AppendLine("- Правый отступ должен быть 0.");
            }
            if (spacing?.Before != null && spacing.Before.Value != "0")
            {
                _errorMessages.AppendLine("- Отступ до абзаца должен быть 0.");
            }
            if (spacing?.After != null && spacing.After.Value != "0")
            {
                _errorMessages.AppendLine("- Отступ после абзаца должен быть 0.");
            }

            return _errorMessages.Length == 0;
        }

        public override bool ValidateRun(Paragraph paragraph, Run run)
        {
            _errorMessages.Clear();

            var styleId = GetStyleId(paragraph);
            if (styleId != "Normal") return true;

            if (run == null)
            {
                _errorMessages.AppendLine("- Текстовый блок (Run) отсутствует.");
            }

            var runProps = run?.RunProperties;
            if (runProps == null)
            {
                _errorMessages.AppendLine("- Свойства текстового блока (Run) отсутствуют.");
            }

            var font = runProps?.RunFonts?.Ascii?.Value;
            var sizeStr = runProps?.FontSize?.Val?.Value;
            if (string.IsNullOrEmpty(sizeStr) || !int.TryParse(sizeStr, out var sizeHalfPt))
            {
                _errorMessages.AppendLine("- Не удалось определить размер шрифта.");
            }
            else
            {
                double sizePt = sizeHalfPt / 2.0;
                if (Math.Abs(sizePt - 13) > 0.1)
                {
                    _errorMessages.AppendLine($"- Размер шрифта: {sizePt:F1} pt, ожидается 13 pt.");
                }
            }

            if (font != "Times New Roman")
            {
                _errorMessages.AppendLine($"- Шрифт: {font ?? "не задан"}, ожидается Times New Roman.");
            }

            return _errorMessages.Length == 0;
        }
    }
}