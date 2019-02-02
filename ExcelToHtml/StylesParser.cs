using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToHtml
{
    public class StylesParser
    {
        private readonly ExcelRange Cell;
        private List<string> Styles;

        public StylesParser(ExcelRange cell)
        {
            Cell = cell;
            Styles = new List<string>();
        }

        /// <summary>
        /// Parse a cell excel styles to css styles
        /// </summary>
        /// <returns>Css styles of a cell</returns>
        public string Parse()
        {
            List<string> styles = new List<string>
            {
                GetBackgroundStyle(),
                GetBorderStyle(),
                GetFontStyle(),
                GetHorizontalAlignmentStyle(),
                GetVerticalAlignmentStyle()
            };
            return string.Join("", styles.Concat(Styles).ToArray());
        }

        /// <summary>
        /// Add a style
        /// </summary>
        /// <param name="style">Css style</param>
        public void Add(string style)
        {
            if (!style.EndsWith(";"))
            {
                style += ";";
            }

            Styles.Add(style);
        }

        /// <summary>
        /// Convert a cell border to css tyle
        /// </summary>
        /// <param name="border">Cell border</param>
        /// <returns>Cell border as css</returns>
        private string BorderToStyle(ExcelBorderItem border)
        {
            string color = Utils.ExcelColorToHex(border.Color);
            ExcelBorderStyle type = border.Style;
            string size = "1px";
            string style = "solid";

            switch (type)
            {
                // Normal
                case ExcelBorderStyle.Thin:
                case ExcelBorderStyle.Thick:
                    size = "1px";
                    break;
                case ExcelBorderStyle.Dashed:
                    style = "dashed";
                    break;
                case ExcelBorderStyle.Dotted:
                    style = "dotted";
                    break;
                case ExcelBorderStyle.Double:
                    style = "double";
                    break;
                case ExcelBorderStyle.None:
                    style = "none";
                    break;
                case ExcelBorderStyle.DashDot:
                    style = "dotted dashed";
                    break;
                case ExcelBorderStyle.DashDotDot:
                    style = "dotted dashed dotted";
                    break;

                // Medium
                case ExcelBorderStyle.Medium:
                    size = "2px";
                    break;
                case ExcelBorderStyle.MediumDashed:
                    size = "2px";
                    style = "dashed";
                    break;
                case ExcelBorderStyle.MediumDashDot:
                    size = "2px";
                    style = "dotted dashed";
                    break;
                case ExcelBorderStyle.MediumDashDotDot:
                    size = "2px";
                    style = "dotted dashed dotted";
                    break;
            }

            return $"{size} {style} {color}";
        }

        /// <summary>
        /// Get border style of a cell
        /// </summary>
        /// <returns>Border css style</returns>
        public string GetBorderStyle()
        {
            StringBuilder result = new StringBuilder();

            Border border = Cell.Style.Border;
            if (border == null)
            {
                return "";
            }

            if (border.Top.Color.Rgb != null)
            {
                result.Append($"border-top: {BorderToStyle(border.Top)};");
            }

            if (border.Bottom.Color.Rgb != null)
            {
                result.Append($"border-bottom: {BorderToStyle(border.Bottom)};");
            }

            if (border.Left.Color.Rgb != null)
            {
                result.Append($"border-left: {BorderToStyle(border.Left)};");
            }

            if (border.Right.Color.Rgb != null)
            {
                result.Append($"border-right: {BorderToStyle(border.Right)};");
            }

            return result.ToString();
        }

        /// <summary>
        /// Get background style of a cell
        /// </summary>
        /// <returns>Background css style</returns>
        public string GetBackgroundStyle()
        {
            ExcelFill fill = Cell.Style.Fill;
            string color = "transparent";

            if (fill.BackgroundColor != null && fill.BackgroundColor.Rgb != null)
            {
                color = Utils.ExcelColorToHex(fill.BackgroundColor);
            }

            return $"background-color: {color};";
        }

        /// <summary>
        /// Get font style of a cell
        /// </summary>
        /// <returns>Font css style</returns>
        public string GetFontStyle()
        {
            ExcelFont font = Cell.Style.Font;
            StringBuilder result = new StringBuilder();

            // Weight
            if (font.Bold)
            {
                result.Append("font-weight: bold;");
            }

            // Decoration
            StringBuilder decoration = new StringBuilder();

            if (font.Strike)
            {
                decoration.Append("line-through");
            }

            if (font.UnderLine)
            {
                decoration.Append("underline");
            }

            if (decoration.Length > 0)
            {
                result.Append($"text-decoration: {string.Join(" ", decoration)};");
            }

            // Style
            if (font.Italic)
            {
                result.Append("font-style: italic;");
            }

            // Size
            result.Append($"font-size: {Utils.FormatNumber(font.Size)}pt;");

            // Color
            if (font.Color.Rgb != null)
            {
                result.Append($"color: {Utils.ExcelColorToHex(font.Color)};");
            }

            // Font
            var fontName = font.Name;
            if (fontName != null && fontName != "")
            {
                if (fontName.Contains(" "))
                {
                    fontName = $"\"{fontName}\"";
                }

                if (fontName != "Arial")
                {
                    fontName += ", Arial";
                }
            }
            else
            {
                fontName = "Arial";
            }

            result.Append($"font-family: {fontName};");

            return result.ToString();
        }

        /// <summary>
        /// Get horizontal alignment style of a cell
        /// </summary>
        /// <returns>Horizontal alignment css style</returns>
        public string GetHorizontalAlignmentStyle()
        {
            ExcelHorizontalAlignment hAlign = Cell.Style.HorizontalAlignment;
            string result = "";

            switch (hAlign)
            {
                case ExcelHorizontalAlignment.Center:
                    result = "center";
                    break;
                case ExcelHorizontalAlignment.Right:
                    result = "right";
                    break;
                case ExcelHorizontalAlignment.Justify:
                    result = "justify";
                    break;
                default:
                    result = "initial";
                    break;
            }

            return $"text-align: {result};";
        }

        /// <summary>
        /// Get vertical alignment style of a cell
        /// </summary>
        /// <returns>Vertical alignment css style</returns>
        public string GetVerticalAlignmentStyle()
        {
            ExcelVerticalAlignment vAlign = Cell.Style.VerticalAlignment;
            string result = "";

            switch (vAlign)
            {
                case ExcelVerticalAlignment.Top:
                    result = "top";
                    break;
                case ExcelVerticalAlignment.Center:
                    result = "middle";
                    break;
                case ExcelVerticalAlignment.Bottom:
                    result = "bottom";
                    break;
                default:
                    result = "initial";
                    break;
            }

            return $"vertical-align: {result};";
        }
    }
}
