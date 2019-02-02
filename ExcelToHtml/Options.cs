namespace ExcelToHtml
{
    public class Options
    {
        public bool Debug { get; set; }
        public bool BeutifyHtml { get; set; }

        public Options()
        {
            Debug = false;
            BeutifyHtml = true;
        }
    }
}
