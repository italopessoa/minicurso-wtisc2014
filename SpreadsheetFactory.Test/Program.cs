using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetFactory.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Header header = new Header();
            header.AddFilter("klkljjkljkl", 10.3);
            header.AddFilter("nome", "italo");
            header.AddFilter("idade", 1);
            header.AddFilter("lkj", DateTime.Now);

            short a = 1;
            header.AddFilter("lkj12e2e2e", a);
            header.Title = "Titulo";
            header.SheetName = "Sheet numero 1";

            SpreadsheetFactory spf = new SpreadsheetFactory();
            spf.Header = header;
            

            WorkbookManager.CreateSpreadsheet(spf);
            WorkbookManager.SaveSpreadsheet(@"E:\source\SVN\pessoal\trunk\SpreadsheetFactory\bin\Debug\", "teste.xls");
        }
    }
}
