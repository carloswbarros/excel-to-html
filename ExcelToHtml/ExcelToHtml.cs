using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using TidyManaged;

namespace ExcelToHtml
{
    public class ExcelToHtml
    {
        private readonly Options options;

        public ExcelToHtml(Options options)
        {
            this.options = options;
        }

        public ExcelToHtml()
        {
            options = new Options();
        }

        /// <summary>
        /// Process file to html string
        /// </summary>
        /// <param name="filePath">File path</param>
        /// <returns>Html string</returns>
        public string Process(string filePath)
        {
            Stream stream = File.Open(filePath, FileMode.Open);
            return Process(stream);
        }

        /// <summary>
        /// Process file stream to html string
        /// </summary>
        /// <param name="stream">Excel file stream</param>
        /// <returns>Html string</returns>
        public string Process(Stream stream)
        {
            ExcelPackage p = new ExcelPackage(stream);

            // Get the first sheet
            var sheet = p.Workbook.Worksheets.FirstOrDefault();
            if (sheet == null)
            {
                throw new Exception("No worksheets found");
            }

            var numberColumns = sheet.Dimension.End.Column;
            var numberRows = sheet.Dimension.End.Row;

            StringBuilder s = new StringBuilder();

            string[] tableStyles = new string[]
            {
                "width: 100%",
                "table-layout: fixed",
                "border-collapse: collapse"
            };

            s.Append($"<table style='{string.Join("; ", tableStyles)}'>");

            // For each row
            for (int i = 1; i <= numberRows; i++)
            {
                ExcelRow row = sheet.Row(i);

                s.Append($"<tr style='height: {Utils.FormatNumber(row.Height)}pt'>");

                // For each column
                for (int j = 1; j <= numberColumns; j++)
                {
                    ExcelColumn column = sheet.Column(j);
                    ExcelRange cell = sheet.Cells[i, j];

                    int colspan = 1;
                    int rowspan = 1;

                    if (!cell.Merge || (cell.Merge && IsFirstMergeRange(sheet, cell.Address, ref colspan, ref rowspan)))
                    {
                        StylesParser styles = new StylesParser(cell);
                        //styles.Add($"width: {Utils.FormatNumber(sheet.Column(j).Width)}pt;");

                        if (options.Debug)
                        {
                            styles.Add("border: 1px solid black;");
                        }

                        s.Append($"<td rowspan='{rowspan}' colspan='{colspan}' style='{styles.Parse()}'>");

                        if (cell.Value != null)
                        {
                            s.Append(cell.Value.ToString());
                        }

                        s.Append("</td>");
                    }
                }

                s.Append("</tr>");
            }

            s.Append("</table>");

            return FormatHtml(s.ToString());
        }

        /// <summary>
        /// Check if is the first cell merged and returns the colspan and rowspan
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="address"></param>
        /// <param name="colspan"></param>
        /// <param name="rowspan"></param>
        /// <returns>If is the first cell merged</returns>
        private bool IsFirstMergeRange(ExcelWorksheet sheet, string address, ref int colspan, ref int rowspan)
        {
            colspan = 1;
            rowspan = 1;
            foreach (var item in sheet.MergedCells)
            {
                var s = item.Split(':');
                if (s.Length > 0 && s[0].Equals(address))
                {

                    ExcelRange range = sheet.Cells[item];
                    colspan = range.End.Column - range.Start.Column;
                    rowspan = range.End.Row - range.Start.Row;
                    if (colspan == 0) colspan = 1;
                    if (rowspan == 0) rowspan = 1;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Format HTML
        /// </summary>
        /// <param name="html">Html string</param>
        /// <returns>Formatted html string</returns>
        private string FormatHtml(string html)
        {
            using (Document doc = Document.FromString(html))
            {
                doc.ShowWarnings = false;
                doc.Quiet = true;
                doc.OutputXhtml = true;
                doc.OutputXml = true;
                doc.IndentAttributes = false;
                doc.AddVerticalSpace = false;
                doc.AddTidyMetaElement = false;

                if (options.BeutifyHtml)
                {
                    doc.WrapAt = 120;
                    doc.TabSize = 4;
                    doc.IndentCdata = true;
                    doc.IndentBlockElements = AutoBool.Yes;
                }

                doc.CleanAndRepair();

                string output = doc.Save();

                if (!options.BeutifyHtml)
                {
                    output = Regex.Replace(output, "\n|\r", "");
                }

                return output;
            }
        }
    }
}
