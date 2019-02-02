using System;
using System.IO;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                ExcelToHtml.ExcelToHtml excelToHtml = new ExcelToHtml.ExcelToHtml(
                    new ExcelToHtml.Options
                    {
                        BeutifyHtml = true,
                        Debug = true
                    }
                );
                string html = excelToHtml.Process("example.xlsx");

                File.WriteAllText("example.html", html);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
            }
        }
    }
}
