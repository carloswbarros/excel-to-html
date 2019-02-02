using OfficeOpenXml.Style;
using System;
using System.Drawing;

namespace ExcelToHtml
{
    public class Utils
    {
        /// <summary>
        /// Format a number to use as a css value
        /// </summary>
        /// <param name="value"></param>
        /// <returns>Formatted number</returns>
        public static string FormatNumber(double value)
        {
            string result = string.Format("{0:0.00}", value);
            return result.Replace(",", ".");
        }

        /// <summary>
        /// Convert a excel color to hex color
        /// </summary>
        /// <param name="color"></param>
        /// <returns>Hex color</returns>
        public static string ExcelColorToHex(ExcelColor color)
        {
            if (color == null)
            {
                return "transparent";
            }

            byte a = (byte)(Convert.ToUInt32(color.Rgb.Substring(0, 2), 16));
            byte r = (byte)(Convert.ToUInt32(color.Rgb.Substring(2, 2), 16));
            byte g = (byte)(Convert.ToUInt32(color.Rgb.Substring(4, 2), 16));
            byte b = (byte)(Convert.ToUInt32(color.Rgb.Substring(6, 2), 16));

            var colorRgb = Color.FromArgb(a, r, g, b);

            return "#" + colorRgb.R.ToString("X2") + colorRgb.G.ToString("X2") + colorRgb.B.ToString("X2");
        }
    }
}
