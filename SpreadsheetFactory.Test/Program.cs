using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using System;
using System.Collections;
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
           
            WorkbookManager WorkbookManager = new global::SpreadsheetFactory.WorkbookManager();

            SpreadsheetFactory spf1 = GerarExemplo(WorkbookManager);
            spf1.MergedTitle = "merged title";
            HSSFCellStyle aas = WorkbookManager.GetNewHSSFCellStyle();
            aas.BorderTop = 1;
            aas.BorderRight = 1;
            aas.BorderRight = 1;
            aas.BorderBottom = 1;
            aas.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            aas.FillForegroundColor = HSSFColor.CORAL.index;
            spf1.SpanTitleStyle = aas;

            SpreadsheetFactory spf2 = GerarExemploSemHeader(WorkbookManager);
            ChildSheet filho = GerarDetalheSemHeaderChild(WorkbookManager);
            filho.FirstCell = 1;
            #region padrao
            //SpreadsheetFactory teste = new SpreadsheetFactory();
            //teste.SpreadsheetFactoryList = new List<SpreadsheetFactory>();
            //spf1.Name = "teste 1";
            //spf1.MergedTitle = "titulo";
            //teste.SpreadsheetFactoryList.Add(spf1);
            //spf2.Name = "teste 1";
            //spf2.MergedTitle = "titulo 1";
            //teste.SpreadsheetFactoryList.Add(spf2);
            #endregion


            #region detalhe

            SpreadsheetFactory teste = new SpreadsheetFactory();
            teste.SpreadsheetFactoryList = new List<SpreadsheetFactory>();
            spf1.Name = "teste 1";
            //spf1.MergedTitle = "titulo";
            
            //seta o objeto que conterá os detalhes
            spf1.ChildSheet = filho;
            //titulo com span
            spf1.ChildSheet.MergedTitle = "Primeiro span";

            //nascimento == "02/02/2013";
            #region linha Amarela


            HSSFCellStyle celulaAmarela0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela0.BorderTop = 1;
            celulaAmarela0.BorderRight = 1;
            celulaAmarela0.BorderRight = 1;
            celulaAmarela0.BorderBottom = 1;

            HSSFCellStyle celulaAmarela1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela1.BorderTop = 1;
            celulaAmarela1.BorderRight = 1;
            celulaAmarela1.BorderRight = 1;
            celulaAmarela1.BorderBottom = 1;

            RowStyle linhaAmarela = new RowStyle();
            linhaAmarela.RowStyleName = "linha amarela2";
            linhaAmarela.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaAmarela.RowForegroundColor = HSSFColor.PINK.index;

            linhaAmarela.AddCellRowStyle(0, celulaAmarela0);
            linhaAmarela.AddCellRowStyle(1, celulaAmarela1);
            //linhaAmarela.AddCellRowStyle(2, celulaAmarela1);
            //linhaAmarela.AddCellRowStyle(3, celulaAmarela1);

            ConditionalFormattingTemplate templateAmarela = new ConditionalFormattingTemplate();
            templateAmarela.Priority = 14;
            templateAmarela.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            templateAmarela.PropertyName = "Nascimento";
            templateAmarela.Value = new DateTime(2013, 2, 2);//"02/02/2013";
            templateAmarela.RowStyle = linhaAmarela;
            spf1.ChildSheet.AddConditionalFormatting("Nascimento", templateAmarela);


            #endregion linha amarela


            //sheet que conterá os valores (usar a mesma da tabela pai)
            spf1.ChildSheet.Name = "teste 1";

            HSSFCellStyle aas2 = WorkbookManager.GetNewHSSFCellStyle();
            aas2.BorderTop = 1;
            aas2.BorderRight = 1;
            aas2.BorderLeft = 1;
            aas2.BorderBottom = 1;
            aas2.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            aas2.FillForegroundColor = HSSFColor.CORNFLOWER_BLUE.index;
            spf1.ChildSheet.SpanTitleStyle = aas2;
            teste.SpreadsheetFactoryList.Add(spf1);

            //seta o objeto que conterá os detalhes
            spf1.ChildSheet.ChildSheet = GerarDetalheSemHeader(WorkbookManager);
            //titulo com span
            spf1.ChildSheet.ChildSheet.MergedTitle = "Segundo espan -  3 NIVEL";
            spf1.ChildSheet.ChildSheet.FirstCell = 2;

            HSSFCellStyle aas3 = WorkbookManager.GetNewHSSFCellStyle();
            aas3.BorderTop = 1;
            aas3.BorderRight = 1;
            aas3.BorderLeft = 1;
            aas3.BorderBottom = 1;
            aas3.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            aas3.FillForegroundColor = HSSFColor.LIME.index;
            spf1.ChildSheet.ChildSheet.SpanTitleStyle = aas3;

            //sheet que conterá os valores (usar a mesma da tabela pai)
            spf1.ChildSheet.ChildSheet.Name = "teste 1";

            spf2.Name = "teste 1";
            spf2.MergedTitle = "segunda tabela";
            teste.SpreadsheetFactoryList.Add(spf2);

            #endregion

            //spf1.ChildSheet.Datasource = list;
           // teste.SpreadsheetFactoryList.Add(spf1);

            WorkbookManager.MountSpreadsheet(teste);
        
            WorkbookManager.SaveSpreadsheet("", DateTime.Now.ToString().Replace("/", "").Replace(":", "").Replace(" ", "") + ".xls");
        }

        static SpreadsheetFactory GerarExemplo(WorkbookManager WorkbookManager)
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

            HSSFDataFormat dateFormat = WorkbookManager.GetNewHSSFDataFormat();
            short dateFormatIndex = dateFormat.GetFormat("DD/MM/YYYY");

            List<object> list = new List<object>();
            list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "Italo", Salario = 2 });
            //list.Add(new Pessoa() { Idade = 1, Nascimento = new DateTime(2013, 2, 2), Nome = "Italo", Salario = 10.3 });
            //list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "Italo", Salario = 20 });
            //list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "Italo", Salario = 10.3 });
            //list.Add(new Pessoa() { Idade = 100, Nascimento = DateTime.Now, Nome = "Italo", Salario = 1 });

            List<object> list2 = new List<object>();
            list2.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "JOAO", Salario = 2 });
            list2.Add(new Pessoa() { Idade = 1, Nascimento = new DateTime(2013, 2, 2), Nome = "JOAO", Salario = 10.3 });
            list2.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "JOAO", Salario = 20 });
            list2.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "JOAO", Salario = 10.3 });
            list2.Add(new Pessoa() { Idade = 100, Nascimento = DateTime.Now, Nome = "JOAO", Salario = 1 });
            (list[0] as Pessoa).Lista = list2;

            List<object> list3 = new List<object>();
            list3.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "JasdadasdasdOAO", Salario = 2 });
            (list2[0] as Pessoa).Lista = list3;

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

           // WorkbookManager.SetHeaderCellStyle(headerStyle);

            HSSFCellStyle contentStyle = WorkbookManager.GetNewHSSFCellStyle();
            contentStyle.BorderTop = 1;
            contentStyle.BorderRight = 1;
            contentStyle.BorderRight = 1;
            contentStyle.BorderBottom = 1;
            contentStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            contentStyle.FillForegroundColor = HSSFColor.YELLOW.index;

            //salario == 2
            #region linhaVermelha

            HSSFCellStyle celulaVermelha0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha0.BorderTop = 1;
            celulaVermelha0.BorderRight = 1;
            celulaVermelha0.BorderRight = 1;
            celulaVermelha0.BorderBottom = 1;

            HSSFCellStyle celulaVermelha1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha1.BorderTop = 1;
            celulaVermelha1.BorderRight = 1;
            celulaVermelha1.BorderRight = 1;
            celulaVermelha1.BorderBottom = 1;
            celulaVermelha1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaVermelha2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha2.BorderTop = 1;
            celulaVermelha2.BorderRight = 1;
            celulaVermelha2.BorderRight = 1;
            celulaVermelha2.BorderBottom = 1;
            celulaVermelha2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaVermelha3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha3.BorderTop = 1;
            celulaVermelha3.BorderRight = 1;
            celulaVermelha3.BorderRight = 1;
            celulaVermelha3.BorderBottom = 1;

            RowStyle linhaVermelha = new RowStyle();
            linhaVermelha.RowStyleName = "linha coral";
            linhaVermelha.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaVermelha.RowForegroundColor = HSSFColor.RED.index;

            linhaVermelha.AddCellRowStyle(0, celulaVermelha0);
            linhaVermelha.AddCellRowStyle(1, celulaVermelha1);
            linhaVermelha.AddCellRowStyle(2, celulaVermelha2);
            linhaVermelha.AddCellRowStyle(3, celulaVermelha3);

            ConditionalFormattingTemplate template = new ConditionalFormattingTemplate();
            template.Priority = 1;
            template.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            template.PropertyName = "Salario";
            template.Value = 2;
            template.RowStyle = linhaVermelha;
            spf.AddConditionalFormatting("Salario", template);

            #endregion linhaVermelha

            //salario > 2
            #region linha azul


            HSSFCellStyle celulaAzul0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul0.BorderTop = 1;
            celulaAzul0.BorderRight = 1;
            celulaAzul0.BorderRight = 1;
            celulaAzul0.BorderBottom = 1;

            HSSFCellStyle celulaAzul1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul1.BorderTop = 1;
            celulaAzul1.BorderRight = 1;
            celulaAzul1.BorderRight = 1;
            celulaAzul1.BorderBottom = 1;
            celulaAzul1.Alignment = HSSFCellStyle.ALIGN_CENTER;
            celulaAzul1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaAzul2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul2.BorderTop = 1;
            celulaAzul2.BorderRight = 1;
            celulaAzul2.BorderRight = 1;
            celulaAzul2.BorderBottom = 1;
            celulaAzul2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaAzul3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul3.BorderTop = 1;
            celulaAzul3.BorderRight = 1;
            celulaAzul3.BorderRight = 1;
            celulaAzul3.BorderBottom = 1;

            RowStyle linhaAzul = new RowStyle();
            linhaAzul.RowStyleName = "linha coral";
            linhaAzul.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaAzul.RowForegroundColor = HSSFColor.BLUE.index;

            linhaAzul.AddCellRowStyle(0, celulaAzul0);
            linhaAzul.AddCellRowStyle(1, celulaAzul1);
            linhaAzul.AddCellRowStyle(2, celulaAzul2);
            linhaAzul.AddCellRowStyle(3, celulaAzul3);

            ConditionalFormattingTemplate templateAzul = new ConditionalFormattingTemplate();
            templateAzul.Priority = 3;
            templateAzul.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.GT;
            templateAzul.PropertyName = "Salario";
            templateAzul.Value = 2;
            templateAzul.RowStyle = linhaAzul;
            spf.AddConditionalFormatting("Salario", templateAzul);


            #endregion linha azul

            //idade == 1
            #region linha verde

            HSSFCellStyle celulaVerde0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde0.BorderTop = 1;
            celulaVerde0.BorderRight = 1;
            celulaVerde0.BorderRight = 1;
            celulaVerde0.BorderBottom = 1;

            HSSFCellStyle celulaVerde1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde1.BorderTop = 1;
            celulaVerde1.BorderRight = 1;
            celulaVerde1.BorderRight = 1;
            celulaVerde1.BorderBottom = 1;
            celulaVerde1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaVerde2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde2.BorderTop = 1;
            celulaVerde2.BorderRight = 1;
            celulaVerde2.BorderRight = 1;
            celulaVerde2.BorderBottom = 1;
            celulaVerde2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaVerde3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde3.BorderTop = 1;
            celulaVerde3.BorderRight = 1;
            celulaVerde3.BorderRight = 1;
            celulaVerde3.BorderBottom = 1;

            RowStyle linhaVerde = new RowStyle();
            linhaVerde.RowStyleName = "linha coral";
            linhaVerde.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaVerde.RowForegroundColor = HSSFColor.GREEN.index;

            linhaVerde.AddCellRowStyle(0, celulaVerde0);
            linhaVerde.AddCellRowStyle(1, celulaVerde1);
            linhaVerde.AddCellRowStyle(2, celulaVerde2);
            linhaVerde.AddCellRowStyle(3, celulaVerde3);

            ConditionalFormattingTemplate templateVerde = new ConditionalFormattingTemplate();
            templateVerde.Priority = 13;
            templateVerde.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            templateVerde.PropertyName = "Idade";
            templateVerde.Value = 1;
            templateVerde.RowStyle = linhaVerde;
            spf.AddConditionalFormatting("Idade", templateVerde);


            #endregion linha verde

            //nascimento == "02/02/2013";
            #region linha Amarela


            HSSFCellStyle celulaAmarela0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela0.BorderTop = 1;
            celulaAmarela0.BorderRight = 1;
            celulaAmarela0.BorderRight = 1;
            celulaAmarela0.BorderBottom = 1;

            HSSFCellStyle celulaAmarela1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela1.BorderTop = 1;
            celulaAmarela1.BorderRight = 1;
            celulaAmarela1.BorderRight = 1;
            celulaAmarela1.BorderBottom = 1;
            celulaAmarela1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaAmarela2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela2.BorderTop = 1;
            celulaAmarela2.BorderRight = 1;
            celulaAmarela2.BorderRight = 1;
            celulaAmarela2.BorderBottom = 1;
            celulaAmarela2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaAmarela3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela3.BorderTop = 1;
            celulaAmarela3.BorderRight = 1;
            celulaAmarela3.BorderRight = 1;
            celulaAmarela3.BorderBottom = 1;

            RowStyle linhaAmarela = new RowStyle();
            linhaAmarela.RowStyleName = "linha amarela";
            linhaAmarela.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaAmarela.RowForegroundColor = HSSFColor.YELLOW.index;

            linhaAmarela.AddCellRowStyle(0, celulaAmarela0);
            linhaAmarela.AddCellRowStyle(1, celulaAmarela1);
            linhaAmarela.AddCellRowStyle(2, celulaAmarela2);
            linhaAmarela.AddCellRowStyle(3, celulaAmarela3);

            ConditionalFormattingTemplate templateAmarela = new ConditionalFormattingTemplate();
            templateAmarela.Priority = 14;
            templateAmarela.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            templateAmarela.PropertyName = "Nascimento";
            templateAmarela.Value = new DateTime(2013, 2, 2);//"02/02/2013";
            templateAmarela.RowStyle = linhaAmarela;
            spf.AddConditionalFormatting("Nascimento", templateAmarela);


            #endregion linha amarela

            RowStyle rs = new RowStyle();

            spf.FirstHeaderCell = 0;
            spf.HeaderCellStyle = headerStyle;
            spf.RowStyle = rs;

            return spf;
        }

        static SpreadsheetFactory GerarExemploSemHeader(WorkbookManager WorkbookManager)
        {
            SpreadsheetFactory spf = new SpreadsheetFactory();

            HSSFDataFormat dateFormat = WorkbookManager.GetNewHSSFDataFormat();
            short dateFormatIndex = dateFormat.GetFormat("DD/MM/YYYY");

            List<object> list = new List<object>();
            list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "joa", Salario = 2 });
            list.Add(new Pessoa() { Idade = 1, Nascimento = new DateTime(2013, 2, 2), Nome = "Italo", Salario = 10.3 });
            list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "joa", Salario = 20 });
            list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "Italo", Salario = 10.3 });
            list.Add(new Pessoa() { Idade = 100, Nascimento = DateTime.Now, Nome = "Italo", Salario = 1 });

            string[] properties = new string[4];
            properties[0] = "Idade";
            properties[1] = "Nascimento";
            properties[2] = "Nome";
            properties[3] = "Salario";

            spf.Datasource = list;
            spf.Properties = properties;

            HSSFCellStyle contentStyle = WorkbookManager.GetNewHSSFCellStyle();
            contentStyle.BorderTop = 1;
            contentStyle.BorderRight = 1;
            contentStyle.BorderRight = 1;
            contentStyle.BorderBottom = 1;
            contentStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            contentStyle.FillForegroundColor = HSSFColor.YELLOW.index;

            //salario == 2
            #region linhaVermelha

            HSSFCellStyle celulaVermelha0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha0.BorderTop = 1;
            celulaVermelha0.BorderRight = 1;
            celulaVermelha0.BorderRight = 1;
            celulaVermelha0.BorderBottom = 1;

            HSSFCellStyle celulaVermelha1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha1.BorderTop = 1;
            celulaVermelha1.BorderRight = 1;
            celulaVermelha1.BorderRight = 1;
            celulaVermelha1.BorderBottom = 1;
            celulaVermelha1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaVermelha2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha2.BorderTop = 1;
            celulaVermelha2.BorderRight = 1;
            celulaVermelha2.BorderRight = 1;
            celulaVermelha2.BorderBottom = 1;
            celulaVermelha2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaVermelha3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha3.BorderTop = 1;
            celulaVermelha3.BorderRight = 1;
            celulaVermelha3.BorderRight = 1;
            celulaVermelha3.BorderBottom = 1;

            RowStyle linhaVermelha = new RowStyle();
            linhaVermelha.RowStyleName = "linha coral";
            linhaVermelha.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaVermelha.RowForegroundColor = HSSFColor.RED.index;

            linhaVermelha.AddCellRowStyle(0, celulaVermelha0);
            linhaVermelha.AddCellRowStyle(1, celulaVermelha1);
            linhaVermelha.AddCellRowStyle(2, celulaVermelha2);
            linhaVermelha.AddCellRowStyle(3, celulaVermelha3);

            ConditionalFormattingTemplate template = new ConditionalFormattingTemplate();
            template.Priority = 1;
            template.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            template.PropertyName = "Salario";
            template.Value = 2;
            template.RowStyle = linhaVermelha;
            spf.AddConditionalFormatting("Salario", template);

            #endregion linhaVermelha

            //salario > 2
            #region linha azul


            HSSFCellStyle celulaAzul0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul0.BorderTop = 1;
            celulaAzul0.BorderRight = 1;
            celulaAzul0.BorderRight = 1;
            celulaAzul0.BorderBottom = 1;

            HSSFCellStyle celulaAzul1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul1.BorderTop = 1;
            celulaAzul1.BorderRight = 1;
            celulaAzul1.BorderRight = 1;
            celulaAzul1.BorderBottom = 1;
            celulaAzul1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaAzul2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul2.BorderTop = 1;
            celulaAzul2.BorderRight = 1;
            celulaAzul2.BorderRight = 1;
            celulaAzul2.BorderBottom = 1;
            celulaAzul2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaAzul3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul3.BorderTop = 1;
            celulaAzul3.BorderRight = 1;
            celulaAzul3.BorderRight = 1;
            celulaAzul3.BorderBottom = 1;

            RowStyle linhaAzul = new RowStyle();
            linhaAzul.RowStyleName = "linha coral";
            linhaAzul.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaAzul.RowForegroundColor = HSSFColor.BLUE.index;

            linhaAzul.AddCellRowStyle(0, celulaAzul0);
            linhaAzul.AddCellRowStyle(1, celulaAzul1);
            linhaAzul.AddCellRowStyle(2, celulaAzul2);
            linhaAzul.AddCellRowStyle(3, celulaAzul3);

            ConditionalFormattingTemplate templateAzul = new ConditionalFormattingTemplate();
            templateAzul.Priority = 3;
            templateAzul.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.GT;
            templateAzul.PropertyName = "Salario";
            templateAzul.Value = 2;
            templateAzul.RowStyle = linhaAzul;
            spf.AddConditionalFormatting("Salario", templateAzul);


            #endregion linha azul

            //idade == 1
            #region linha verde


            HSSFCellStyle celulaVerde0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde0.BorderTop = 1;
            celulaVerde0.BorderRight = 1;
            celulaVerde0.BorderRight = 1;
            celulaVerde0.BorderBottom = 1;

            HSSFCellStyle celulaVerde1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde1.BorderTop = 1;
            celulaVerde1.BorderRight = 1;
            celulaVerde1.BorderRight = 1;
            celulaVerde1.BorderBottom = 1;
            celulaVerde1.Alignment = HSSFCellStyle.ALIGN_CENTER;
            celulaVerde1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaVerde2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde2.BorderTop = 1;
            celulaVerde2.BorderRight = 1;
            celulaVerde2.BorderRight = 1;
            celulaVerde2.BorderBottom = 1;
            celulaVerde2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaVerde3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde3.BorderTop = 1;
            celulaVerde3.BorderRight = 1;
            celulaVerde3.BorderRight = 1;
            celulaVerde3.BorderBottom = 1;

            RowStyle linhaVerde = new RowStyle();
            linhaVerde.RowStyleName = "linha coral";
            linhaVerde.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaVerde.RowForegroundColor = HSSFColor.GREEN.index;

            linhaVerde.AddCellRowStyle(0, celulaVerde0);
            linhaVerde.AddCellRowStyle(1, celulaVerde1);
            linhaVerde.AddCellRowStyle(2, celulaVerde2);
            linhaVerde.AddCellRowStyle(3, celulaVerde3);

            ConditionalFormattingTemplate templateVerde = new ConditionalFormattingTemplate();
            templateVerde.Priority = 15;
            templateVerde.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            templateVerde.PropertyName = "Idade";
            templateVerde.Value = 1;
            templateVerde.RowStyle = linhaVerde;
            spf.AddConditionalFormatting("Idade", templateVerde);


            #endregion linha verde

            //nascimento == "02/02/2013";
            #region linha Amarela


            HSSFCellStyle celulaAmarela0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela0.BorderTop = 1;
            celulaAmarela0.BorderRight = 1;
            celulaAmarela0.BorderRight = 1;
            celulaAmarela0.BorderBottom = 1;

            HSSFCellStyle celulaAmarela1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela1.BorderTop = 1;
            celulaAmarela1.BorderRight = 1;
            celulaAmarela1.BorderRight = 1;
            celulaAmarela1.BorderBottom = 1;
            celulaAmarela1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaAmarela2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela2.BorderTop = 1;
            celulaAmarela2.BorderRight = 1;
            celulaAmarela2.BorderRight = 1;
            celulaAmarela2.BorderBottom = 1;
            celulaAmarela2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaAmarela3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela3.BorderTop = 1;
            celulaAmarela3.BorderRight = 1;
            celulaAmarela3.BorderRight = 1;
            celulaAmarela3.BorderBottom = 1;

            RowStyle linhaAmarela = new RowStyle();
            linhaAmarela.RowStyleName = "linha amarela";
            linhaAmarela.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaAmarela.RowForegroundColor = HSSFColor.YELLOW.index;

            linhaAmarela.AddCellRowStyle(0, celulaAmarela0);
            linhaAmarela.AddCellRowStyle(1, celulaAmarela1);
            linhaAmarela.AddCellRowStyle(2, celulaAmarela2);
            linhaAmarela.AddCellRowStyle(3, celulaAmarela3);

            ConditionalFormattingTemplate templateAmarela = new ConditionalFormattingTemplate();
            templateAmarela.Priority = 14;
            templateAmarela.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            templateAmarela.PropertyName = "Nascimento";
            templateAmarela.Value = new DateTime(2013, 2, 2);//"02/02/2013";
            templateAmarela.RowStyle = linhaAmarela;
            spf.AddConditionalFormatting("Nascimento", templateAmarela);


            #endregion linha amarela

            RowStyle rs = new RowStyle();

            spf.FirstHeaderCell = 0;
            spf.RowStyle = rs;
            return spf;
        }


        static ChildSheet GerarDetalheSemHeaderChild(WorkbookManager WorkbookManager)
        {
            ChildSheet spf = new ChildSheet("Lista");

            HSSFDataFormat dateFormat = WorkbookManager.GetNewHSSFDataFormat();
            short dateFormatIndex = dateFormat.GetFormat("DD/MM/YYYY");

            string[] properties = new string[2];
            properties[0] = "Idade";
            properties[1] = "Salario";

            spf.Properties = properties;

            HSSFCellStyle contentStyle = WorkbookManager.GetNewHSSFCellStyle();
            contentStyle.BorderTop = 1;
            contentStyle.BorderRight = 1;
            contentStyle.BorderRight = 1;
            contentStyle.BorderBottom = 1;
            contentStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            contentStyle.FillForegroundColor = HSSFColor.YELLOW.index;

            //salario == 2
            #region linhaVermelha

            HSSFCellStyle celulaVermelha0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha0.BorderTop = 1;
            celulaVermelha0.BorderRight = 1;
            celulaVermelha0.BorderRight = 1;
            celulaVermelha0.BorderBottom = 1;

            HSSFCellStyle celulaVermelha1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha1.BorderTop = 1;
            celulaVermelha1.BorderRight = 1;
            celulaVermelha1.BorderRight = 1;
            celulaVermelha1.BorderBottom = 1;
            celulaVermelha1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaVermelha2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha2.BorderTop = 1;
            celulaVermelha2.BorderRight = 1;
            celulaVermelha2.BorderRight = 1;
            celulaVermelha2.BorderBottom = 1;
            celulaVermelha2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaVermelha3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha3.BorderTop = 1;
            celulaVermelha3.BorderRight = 1;
            celulaVermelha3.BorderRight = 1;
            celulaVermelha3.BorderBottom = 1;

            RowStyle linhaVermelha = new RowStyle();
            linhaVermelha.RowStyleName = "linha coral";
            linhaVermelha.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaVermelha.RowForegroundColor = HSSFColor.RED.index;

            linhaVermelha.AddCellRowStyle(0, celulaVermelha0);
            linhaVermelha.AddCellRowStyle(1, celulaVermelha1);

            ConditionalFormattingTemplate template = new ConditionalFormattingTemplate();
            template.Priority = 1;
            template.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            template.PropertyName = "Salario";
            template.Value = 2;
            template.RowStyle = linhaVermelha;
            spf.AddConditionalFormatting("Salario", template);

            #endregion linhaVermelha

            //salario > 2
            #region linha azul


            HSSFCellStyle celulaAzul0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul0.BorderTop = 1;
            celulaAzul0.BorderRight = 1;
            celulaAzul0.BorderRight = 1;
            celulaAzul0.BorderBottom = 1;

            HSSFCellStyle celulaAzul1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul1.BorderTop = 1;
            celulaAzul1.BorderRight = 1;
            celulaAzul1.BorderRight = 1;
            celulaAzul1.BorderBottom = 1;
            celulaAzul1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaAzul2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul2.BorderTop = 1;
            celulaAzul2.BorderRight = 1;
            celulaAzul2.BorderRight = 1;
            celulaAzul2.BorderBottom = 1;
            celulaAzul2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaAzul3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul3.BorderTop = 1;
            celulaAzul3.BorderRight = 1;
            celulaAzul3.BorderRight = 1;
            celulaAzul3.BorderBottom = 1;

            RowStyle linhaAzul = new RowStyle();
            linhaAzul.RowStyleName = "linha coral";
            linhaAzul.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaAzul.RowForegroundColor = HSSFColor.BLUE.index;

            linhaAzul.AddCellRowStyle(0, celulaAzul0);
            linhaAzul.AddCellRowStyle(1, celulaAzul1);

            ConditionalFormattingTemplate templateAzul = new ConditionalFormattingTemplate();
            templateAzul.Priority = 3;
            templateAzul.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.GT;
            templateAzul.PropertyName = "Salario";
            templateAzul.Value = 2;
            templateAzul.RowStyle = linhaAzul;
            spf.AddConditionalFormatting("Salario", templateAzul);


            #endregion linha azul

            //idade == 1
            #region linha verde


            HSSFCellStyle celulaVerde0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde0.BorderTop = 1;
            celulaVerde0.BorderRight = 1;
            celulaVerde0.BorderRight = 1;
            celulaVerde0.BorderBottom = 1;

            HSSFCellStyle celulaVerde1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde1.BorderTop = 1;
            celulaVerde1.BorderRight = 1;
            celulaVerde1.BorderRight = 1;
            celulaVerde1.BorderBottom = 1;
            celulaVerde1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaVerde2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde2.BorderTop = 1;
            celulaVerde2.BorderRight = 1;
            celulaVerde2.BorderRight = 1;
            celulaVerde2.BorderBottom = 1;
            celulaVerde2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaVerde3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde3.BorderTop = 1;
            celulaVerde3.BorderRight = 1;
            celulaVerde3.BorderRight = 1;
            celulaVerde3.BorderBottom = 1;

            RowStyle linhaVerde = new RowStyle();
            linhaVerde.RowStyleName = "linha coral";
            linhaVerde.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaVerde.RowForegroundColor = HSSFColor.GREEN.index;

            linhaVerde.AddCellRowStyle(0, celulaVerde0);
            linhaVerde.AddCellRowStyle(1, celulaVerde1);

            ConditionalFormattingTemplate templateVerde = new ConditionalFormattingTemplate();
            templateVerde.Priority = 13;
            templateVerde.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            templateVerde.PropertyName = "Idade";
            templateVerde.Value = 1;
            templateVerde.RowStyle = linhaVerde;
            spf.AddConditionalFormatting("Idade", templateVerde);


            #endregion linha verde

            //nascimento == "02/02/2013";
            #region linha Amarela


            HSSFCellStyle celulaAmarela0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela0.BorderTop = 1;
            celulaAmarela0.BorderRight = 1;
            celulaAmarela0.BorderRight = 1;
            celulaAmarela0.BorderBottom = 1;

            HSSFCellStyle celulaAmarela1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela1.BorderTop = 1;
            celulaAmarela1.BorderRight = 1;
            celulaAmarela1.BorderRight = 1;
            celulaAmarela1.BorderBottom = 1;
            celulaAmarela1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaAmarela2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela2.BorderTop = 1;
            celulaAmarela2.BorderRight = 1;
            celulaAmarela2.BorderRight = 1;
            celulaAmarela2.BorderBottom = 1;
            celulaAmarela2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaAmarela3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela3.BorderTop = 1;
            celulaAmarela3.BorderRight = 1;
            celulaAmarela3.BorderRight = 1;
            celulaAmarela3.BorderBottom = 1;

            RowStyle linhaAmarela = new RowStyle();
            linhaAmarela.RowStyleName = "linha amarela";
            linhaAmarela.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaAmarela.RowForegroundColor = HSSFColor.YELLOW.index;

            linhaAmarela.AddCellRowStyle(0, celulaAmarela0);
            linhaAmarela.AddCellRowStyle(1, celulaAmarela1);

            ConditionalFormattingTemplate templateAmarela = new ConditionalFormattingTemplate();
            templateAmarela.Priority = 14;
            templateAmarela.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            templateAmarela.PropertyName = "Nascimento";
            templateAmarela.Value = new DateTime(2013, 2, 2);//"02/02/2013";
            templateAmarela.RowStyle = linhaAmarela;
            //spf.AddConditionalFormatting("Nascimento", templateAmarela);


            #endregion linha amarela

            RowStyle rs = new RowStyle();

            spf.FirstHeaderCell = 0;
            spf.RowStyle = rs;
            return spf;
        }

        static ChildSheet GerarDetalheSemHeader(WorkbookManager WorkbookManager)
        {
            ChildSheet spf = new ChildSheet("Lista");

            HSSFDataFormat dateFormat = WorkbookManager.GetNewHSSFDataFormat();
            short dateFormatIndex = dateFormat.GetFormat("DD/MM/YYYY");

            string[] properties = new string[2];
            properties[0] = "Idade";
            properties[1] = "Salario";

            spf.Properties = properties;

            HSSFCellStyle contentStyle = WorkbookManager.GetNewHSSFCellStyle();
            contentStyle.BorderTop = 1;
            contentStyle.BorderRight = 1;
            contentStyle.BorderRight = 1;
            contentStyle.BorderBottom = 1;
            contentStyle.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            contentStyle.FillForegroundColor = HSSFColor.YELLOW.index;

            //salario == 2
            #region linhaVermelha

            HSSFCellStyle celulaVermelha0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha0.BorderTop = 1;
            celulaVermelha0.BorderRight = 1;
            celulaVermelha0.BorderRight = 1;
            celulaVermelha0.BorderBottom = 1;

            HSSFCellStyle celulaVermelha1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha1.BorderTop = 1;
            celulaVermelha1.BorderRight = 1;
            celulaVermelha1.BorderRight = 1;
            celulaVermelha1.BorderBottom = 1;
            celulaVermelha1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaVermelha2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha2.BorderTop = 1;
            celulaVermelha2.BorderRight = 1;
            celulaVermelha2.BorderRight = 1;
            celulaVermelha2.BorderBottom = 1;
            celulaVermelha2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaVermelha3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVermelha3.BorderTop = 1;
            celulaVermelha3.BorderRight = 1;
            celulaVermelha3.BorderRight = 1;
            celulaVermelha3.BorderBottom = 1;

            RowStyle linhaVermelha = new RowStyle();
            linhaVermelha.RowStyleName = "linha coral";
            linhaVermelha.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaVermelha.RowForegroundColor = HSSFColor.RED.index;

            linhaVermelha.AddCellRowStyle(0, celulaVermelha0);
            linhaVermelha.AddCellRowStyle(1, celulaVermelha1);
            linhaVermelha.AddCellRowStyle(2, celulaVermelha2);
            linhaVermelha.AddCellRowStyle(3, celulaVermelha3);

            ConditionalFormattingTemplate template = new ConditionalFormattingTemplate();
            template.Priority = 1;
            template.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            template.PropertyName = "Salario";
            template.Value = 2;
            template.RowStyle = linhaVermelha;
            spf.AddConditionalFormatting("Salario", template);

            #endregion linhaVermelha

            //salario > 2
            #region linha azul


            HSSFCellStyle celulaAzul0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul0.BorderTop = 1;
            celulaAzul0.BorderRight = 1;
            celulaAzul0.BorderRight = 1;
            celulaAzul0.BorderBottom = 1;

            HSSFCellStyle celulaAzul1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul1.BorderTop = 1;
            celulaAzul1.BorderRight = 1;
            celulaAzul1.BorderRight = 1;
            celulaAzul1.BorderBottom = 1;
            celulaAzul1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaAzul2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul2.BorderTop = 1;
            celulaAzul2.BorderRight = 1;
            celulaAzul2.BorderRight = 1;
            celulaAzul2.BorderBottom = 1;
            celulaAzul2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaAzul3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAzul3.BorderTop = 1;
            celulaAzul3.BorderRight = 1;
            celulaAzul3.BorderRight = 1;
            celulaAzul3.BorderBottom = 1;

            RowStyle linhaAzul = new RowStyle();
            linhaAzul.RowStyleName = "linha coral";
            linhaAzul.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaAzul.RowForegroundColor = HSSFColor.BLUE.index;

            linhaAzul.AddCellRowStyle(0, celulaAzul0);
            linhaAzul.AddCellRowStyle(1, celulaAzul1);
            linhaAzul.AddCellRowStyle(2, celulaAzul2);
            linhaAzul.AddCellRowStyle(3, celulaAzul3);

            ConditionalFormattingTemplate templateAzul = new ConditionalFormattingTemplate();
            templateAzul.Priority = 3;
            templateAzul.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.GT;
            templateAzul.PropertyName = "Salario";
            templateAzul.Value = 2;
            templateAzul.RowStyle = linhaAzul;
            spf.AddConditionalFormatting("Salario", templateAzul);


            #endregion linha azul

            //idade == 1
            #region linha verde


            HSSFCellStyle celulaVerde0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde0.BorderTop = 1;
            celulaVerde0.BorderRight = 1;
            celulaVerde0.BorderRight = 1;
            celulaVerde0.BorderBottom = 1;

            HSSFCellStyle celulaVerde1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde1.BorderTop = 1;
            celulaVerde1.BorderRight = 1;
            celulaVerde1.BorderRight = 1;
            celulaVerde1.BorderBottom = 1;
            celulaVerde1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaVerde2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde2.BorderTop = 1;
            celulaVerde2.BorderRight = 1;
            celulaVerde2.BorderRight = 1;
            celulaVerde2.BorderBottom = 1;
            celulaVerde2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaVerde3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaVerde3.BorderTop = 1;
            celulaVerde3.BorderRight = 1;
            celulaVerde3.BorderRight = 1;
            celulaVerde3.BorderBottom = 1;

            RowStyle linhaVerde = new RowStyle();
            linhaVerde.RowStyleName = "linha coral";
            linhaVerde.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaVerde.RowForegroundColor = HSSFColor.GREEN.index;

            linhaVerde.AddCellRowStyle(0, celulaVerde0);
            linhaVerde.AddCellRowStyle(1, celulaVerde1);
            linhaVerde.AddCellRowStyle(2, celulaVerde2);
            linhaVerde.AddCellRowStyle(3, celulaVerde3);

            ConditionalFormattingTemplate templateVerde = new ConditionalFormattingTemplate();
            templateVerde.Priority = 13;
            templateVerde.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            templateVerde.PropertyName = "Idade";
            templateVerde.Value = 1;
            templateVerde.RowStyle = linhaVerde;
            spf.AddConditionalFormatting("Idade", templateVerde);


            #endregion linha verde

            //nascimento == "02/02/2013";
            #region linha Amarela


            HSSFCellStyle celulaAmarela0 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela0.BorderTop = 1;
            celulaAmarela0.BorderRight = 1;
            celulaAmarela0.BorderRight = 1;
            celulaAmarela0.BorderBottom = 1;

            HSSFCellStyle celulaAmarela1 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela1.BorderTop = 1;
            celulaAmarela1.BorderRight = 1;
            celulaAmarela1.BorderRight = 1;
            celulaAmarela1.BorderBottom = 1;
            celulaAmarela1.DataFormat = dateFormatIndex;

            HSSFCellStyle celulaAmarela2 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela2.BorderTop = 1;
            celulaAmarela2.BorderRight = 1;
            celulaAmarela2.BorderRight = 1;
            celulaAmarela2.BorderBottom = 1;
            celulaAmarela2.Alignment = HSSFCellStyle.ALIGN_CENTER;

            HSSFCellStyle celulaAmarela3 = WorkbookManager.GetNewHSSFCellStyle();
            celulaAmarela3.BorderTop = 1;
            celulaAmarela3.BorderRight = 1;
            celulaAmarela3.BorderRight = 1;
            celulaAmarela3.BorderBottom = 1;

            RowStyle linhaAmarela = new RowStyle();
            linhaAmarela.RowStyleName = "linha amarela";
            linhaAmarela.RowFillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            linhaAmarela.RowForegroundColor = HSSFColor.YELLOW.index;

            linhaAmarela.AddCellRowStyle(0, celulaAmarela0);
            linhaAmarela.AddCellRowStyle(1, celulaAmarela1);
            linhaAmarela.AddCellRowStyle(2, celulaAmarela2);
            linhaAmarela.AddCellRowStyle(3, celulaAmarela3);

            ConditionalFormattingTemplate templateAmarela = new ConditionalFormattingTemplate();
            templateAmarela.Priority = 14;
            templateAmarela.ComparisonOperator = NPOI.HSSF.Record.ComparisonOperator.EQUAL;
            templateAmarela.PropertyName = "Nascimento";
            templateAmarela.Value = new DateTime(2013, 2, 2);//"02/02/2013";
            templateAmarela.RowStyle = linhaAmarela;
            //spf.AddConditionalFormatting("Nascimento", templateAmarela);


            #endregion linha amarela

            RowStyle rs = new RowStyle();
            
            spf.FirstHeaderCell = 0;
            spf.RowStyle = rs;
            return spf;
        }

        static void Ma3in(string[] args)
        {
            //HSSFWorkbook wb = new HSSFWorkbook(); //or new HSSFWorkbook();

            //CreationHelper factory = wb.cre getCreationHelper();

            //Sheet sheet = wb.createSheet();

            //Rowl row  = sheet.createRow(3);
            //Cell cell = row.createCell(5);
            //cell.setCellValue("F4");

            //Drawing drawing = sheet.createDrawingPatriarch();

            //// When the comment box is visible, have it show in a 1x3 space
            //ClientAnchor anchor = factory.createClientAnchor();
            //anchor.setCol1(cell.getColumnIndex());
            //anchor.setCol2(cell.getColumnIndex()+1);
            //anchor.setRow1(row.getRowNul());
            //anchor.setRow2(row.getRowNul()+3);

            //// Create the comment and set the text+author
            //Comment comment = drawing.createCellComment(anchor);
            //RichTextString str = factory.createRichTextString("Hello, World!");
            //comment.setString(str);
            //comment.setAuthor("Apache POI");

            //// Assign the comment to the cell
            //cell.setCellComment(comment);

            //String fname = "comment-xssf.xls";
            //if(wb instanceof XSSFWorkbook) fname += "x";
            //FileOutputStream out = new FileOutputStream(fname);
            //wb.write(out);
            //out.close();

            IDictionary<string, List<ConditionalFormattingTemplate>> b = new Dictionary<string, List<ConditionalFormattingTemplate>>();
            b["aaq"].Sort(delegate(ConditionalFormattingTemplate a, ConditionalFormattingTemplate i)
            { return a.Priority.CompareTo(i.Priority); });

            ConditionalFormattingTemplate de = new ConditionalFormattingTemplate();

        }

        static void Maiwn(string[] args)
        {
            A a = new A();
            a.Idade = 1;
            a.MyProperty = new B();
            a.MyProperty.Idade = new List<object>();


            a.MyProperty.MyPropertyA = a;

            Console.WriteLine(GetPropValue(a, "MyProperty.MyPropertyA.Idade"));
            Console.Read();
        }

        public static object GetPropValue(object src, string propName)
        {
            //TODO: é realmente necessário utilizar uma lista generica de object?
            Type list = typeof(System.Collections.Generic.IList<object>);
            string listNamespace = list.Namespace;

            object aux = src;
            string[] nivels = propName.Split('.');

            object property = null;

            foreach (var nivel in nivels)
            {
                property = aux.GetType().GetProperty(nivel).GetValue(aux, null);
                if (property != null)
                {
                    aux = property;
                }
            }

            if (property!= null && !property.GetType().Namespace.Equals(listNamespace))
            {
                //TODO: melhorar a mensagem informando a propriedade
                throw new ArgumentException("o objeto indicado  não é uma lista");
            }

            return property;
        }

        internal class A
        {
            public B MyProperty { get; set; }
            public int Idade { get; set; }
        }

        internal class B
        {
            public IList<object> MyProperty { get; set; }
            public IList<object> Idade { get; set; }
            public StringBuilder Data { get; set; }
            public A MyPropertyA { get; set; }
        }
    }
}

