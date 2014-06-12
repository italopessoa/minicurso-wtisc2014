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
            spf.TableHeaders = new List<TableHeader>();

            #region table header


            TableHeader ta = new TableHeader();
            ta.Text = "A";

            TableHeader tb = new TableHeader();
            tb.Text = "B";
            tb.AddSpanCell("1");
            tb.AddSpanCell("2");

            TableHeader tc = new TableHeader();
            tc.Text = "C";
            
            TableHeader td = new TableHeader();
            td.Text = "D";
            td.AddSpanCell("3");
            td.AddSpanCell("4");

            TableHeader te = new TableHeader();
            te.Text = "E";
            te.AddSpanCell("5");
            te.AddSpanCell("6");

            tc.Cells = new List<TableHeader>();
            tc.Cells.Add(td);
            tc.Cells.Add(te);

            //spf.TableHeaders.Add(tx);
            //spf.TableHeaders.Add(ta);
            //spf.TableHeaders.Add(tb);
            spf.TableHeaders.Add(tc);
            //spf.TableHeaders.Add(td);
            //spf.TableHeaders.Add(te);

            #endregion table header

            WorkbookManager.CreateSpreadsheet(spf);
            WorkbookManager.SaveSpreadsheet(@"E:\source\SVN\pessoal\trunk\SpreadsheetFactory\bin\Debug\", DateTime.Now.Second+DateTime.Now.Millisecond+"teste.xls");
        }
    }
}
