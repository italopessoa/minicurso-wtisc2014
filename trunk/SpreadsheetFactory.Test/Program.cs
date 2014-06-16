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
            header.AddFilter("Valor", 10.3);
            header.AddFilter("Nome", "italo");
            header.AddFilter("Idade", 22);
            header.AddFilter("Data de nascimento", DateTime.Now);

            short a = 1;
            header.Title = "Testando exportação";
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

            spf.TableHeaders.Add(ta);
            spf.TableHeaders.Add(tb);
            spf.TableHeaders.Add(tc);

            #endregion table header


            #region teste final


            //TableHeader ta = new TableHeader();
            //ta.Text = "A";

            //TableHeader tb = new TableHeader();
            //tb.Text = "B";
            //tb.AddSpanCell("1");
            //tb.AddSpanCell("2");

            //TableHeader tc = new TableHeader();
            //tc.Text = "C";

            //TableHeader td = new TableHeader();
            //td.Text = "D";

            //TableHeader te = new TableHeader();
            //te.Text = "E";

            //TableHeader tf = new TableHeader();
            //tf.Text = "F";
            //tf.AddSpanCell("3");
            //tf.AddSpanCell("4");

            //TableHeader tg = new TableHeader();
            //tg.Text = "G";
            //tg.AddSpanCell("5");
            //tg.AddSpanCell("6");

            //TableHeader th = new TableHeader();
            //th.Text = "H";
            //th.AddSpanCell("7");
            //th.AddSpanCell("8");

            //TableHeader ti = new TableHeader();
            //ti.Text = "I";
            //ti.AddSpanCell("9");
            //ti.AddSpanCell("10");

            //td.Cells = new List<TableHeader>();
            //td.Cells.Add(tf);
            //td.Cells.Add(tg);

            //te.Cells = new List<TableHeader>();
            //te.Cells.Add(th);
            //te.Cells.Add(ti);

            //tc.Cells = new List<TableHeader>();
            //tc.Cells.Add(td);
            //tc.Cells.Add(te);

            //spf.TableHeaders.Add(ta);
            //spf.TableHeaders.Add(tb);
            //spf.TableHeaders.Add(tc);

            #endregion


            WorkbookManager.CreateSpreadsheet(spf);
            //Console.WriteLine("FIM");
            //Console.ReadKey();
            WorkbookManager.SaveSpreadsheet("", DateTime.Now.ToString().Replace("/", "").Replace(":","").Replace(" ","")+".xls");
        }
    }

   
}
