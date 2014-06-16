using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
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

            #region teste brutal

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

            //TableHeader tg = new TableHeader();
            //tg.Text = "G";

            //TableHeader th = new TableHeader();
            //th.Text = "H";

            //TableHeader ti = new TableHeader();
            //ti.Text = "I";

            //TableHeader tj = new TableHeader();
            //tj.Text = "J";
            //tj.AddSpanCell("3");
            //tj.AddSpanCell("4");

            //TableHeader tk = new TableHeader();
            //tk.Text = "K";
            //tk.AddSpanCell("5");
            //tk.AddSpanCell("6");

            //TableHeader tl = new TableHeader();
            //tl.Text = "L";
            //tl.AddSpanCell("7");
            //tl.AddSpanCell("8");

            //TableHeader tm = new TableHeader();
            //tm.Text = "M";
            //tm.AddSpanCell("9");
            //tm.AddSpanCell("10");

            //TableHeader tn = new TableHeader();
            //tn.Text = "N";
            //tn.AddSpanCell("11");
            //tn.AddSpanCell("12");

            //TableHeader to = new TableHeader();
            //to.Text = "O";
            //to.AddSpanCell("13");
            //to.AddSpanCell("14");

            //TableHeader tp = new TableHeader();
            //tp.Text = "P";
            //tp.AddSpanCell("15");
            //tp.AddSpanCell("16");

            //TableHeader tq = new TableHeader();
            //tq.Text = "Q";
            //tq.AddSpanCell("17");
            //tq.AddSpanCell("18");


            //#region tc

            //tc.Cells = new List<TableHeader>();
            //td.Cells = new List<TableHeader>();
            //te.Cells = new List<TableHeader>();
            //tf.Cells = new List<TableHeader>();
            //tg.Cells = new List<TableHeader>();
            //th.Cells = new List<TableHeader>();
            //ti.Cells = new List<TableHeader>();

            //#region td

            //tf.Cells.Add(tj);
            //tf.Cells.Add(tk);

            //tg.Cells.Add(tl);
            //tg.Cells.Add(tm);

            //td.Cells.Add(tf);
            //td.Cells.Add(tg);

            //#endregion td

            //#region te

            //th.Cells.Add(tn);
            //th.Cells.Add(to);

            //ti.Cells.Add(tp);
            //ti.Cells.Add(tq);

            //te.Cells.Add(th);
            //te.Cells.Add(ti);

            //#endregion te


            //tc.Cells.Add(td);
            //tc.Cells.Add(te);
            //#endregion tc

            //spf.TableHeaders.Add(ta);
            //spf.TableHeaders.Add(tb);
            //spf.TableHeaders.Add(tc);

            #endregion

            List<object> list = new List<object>();
            list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "Italo", Salario = 2 });
            list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "Italo", Salario = 10.3 });
            list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "Italo", Salario = 10.3 });
            list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "Italo", Salario = 10.3 });

            string[] properties = new string[4];
            properties[0] = "Idade";
            properties[1] = "Nascimento";
            properties[2] = "Nome";
            properties[3] = "Salario";

            spf.Datasource = list;
            spf.Properties = properties;


            HSSFCellStyle headerStyle = WorkbookManager.GetNewHSSFCellStyle();
            headerStyle.BorderTop = 1;
            headerStyle.BorderRight = 1;
            headerStyle.BorderRight = 1;
            headerStyle.BorderBottom = 1;

            headerStyle.Alignment = HSSFCellStyle.ALIGN_CENTER;
            headerStyle.VerticalAlignment = HSSFCellStyle.VERTICAL_CENTER;
            headerStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            headerStyle.FillForegroundColor = HSSFColor.BLUE.index;

            HSSFFont font = WorkbookManager.GetNewHSSFCellFont();
            font.Color = HSSFColor.WHITE.index;
            font.Boldweight = HSSFFont.BOLDWEIGHT_BOLD;
            headerStyle.SetFont(font);

            WorkbookManager.SetHeaderCellStyle(headerStyle);

            HSSFCellStyle contentStyle = WorkbookManager.GetNewHSSFCellStyle();
            contentStyle.BorderTop = 1;
            contentStyle.BorderRight = 1;
            contentStyle.BorderRight = 1;
            contentStyle.BorderBottom = 1;
            contentStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            contentStyle.FillForegroundColor = HSSFColor.YELLOW.index;

            WorkbookManager.SetDefaultContentCellStyle(contentStyle);

            WorkbookManager.CreateSpreadsheet(spf);
            //Console.WriteLine("FIM");
            //Console.ReadKey();
            WorkbookManager.SaveSpreadsheet("", DateTime.Now.ToString().Replace("/", "").Replace(":","").Replace(" ","")+".xls");
        }
    }

   
}

