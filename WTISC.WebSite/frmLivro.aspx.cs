using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using WTISC._2014.Data;
using WTISC._2014.Data.Exceptions;
using WTISC._2014.Data.Management;

public partial class frmLivro : System.Web.UI.Page
{
    #region membros privados

    private MGenero mGenero;
    private MAutor mAutor;
    private MLivro mLivro;
    private const string TIPO_CONSULTA_LIVROS = "tipo_consulta_livros";

    #endregion

    protected void Page_Load(object sender, EventArgs e)
    {
        this.mGenero = new MGenero();
        this.mAutor = new MAutor();
        this.mLivro = new MLivro();
        this.CarregarGeneros();
        this.CarregarAutores();

        string status = Request.QueryString["status"];

        if (!String.IsNullOrEmpty(status))
        {
            if (status.Equals("lido"))
            {
                Session[TIPO_CONSULTA_LIVROS] = true;
                this.CarregarLivros(true);
            }
            else if (status.Equals("nlido"))
            {
                Session[TIPO_CONSULTA_LIVROS] = false;
                this.CarregarLivros(false);
            }
        }
        else
        {
            Session[TIPO_CONSULTA_LIVROS] = null;
            this.CarregarLivros(null);
        }
    }

    #region metodos privados

    private void CarregarGeneros()
    {
        this.ddlGeneroLivro.DataValueField = "Id";
        this.ddlGeneroLivro.DataTextField = "Nome";
        this.ddlGeneroLivro.DataSource = this.mGenero.FindAll();
        this.ddlGeneroLivro.DataBind();
    }

    private void CarregarAutores()
    {
        this.ddlAutorLivro.DataValueField = "Id";
        this.ddlAutorLivro.DataTextField = "Nome";
        this.ddlAutorLivro.DataSource = this.mAutor.FindAll();
        this.ddlAutorLivro.DataBind();
    }

    private void CarregarLivros(bool? lido)
    {
        if (lido.HasValue)
        {
            this.pnlCadastroGenero.Visible = false;
            this.gdvLivros.DataSource = this.mLivro.FindAll().FindAll(l => l.Lido == lido);
            this.gdvLivros.DataBind();
        }
        else
        {
            this.gdvLivros.DataSource = this.mLivro.FindAll();
            this.gdvLivros.DataBind();
        }
    }

    private void LimparTela()
    {
        this.btnAlterarLivro.Visible = false;
        this.btnCadLivro.Visible = true;
        this.btnCancelarEdicao.Visible = false;

        this.txtIdLivro.Text = String.Empty;
        this.txtTituloLivro.Text = string.Empty;
        this.txtDescricaoLivro.Text = string.Empty;
        this.cbLivroLido.Checked = false;
    }

    private void RemoverLivro(int isbn)
    {
        this.mLivro.DeleteBook(isbn);
    }

    #endregion


    #region gridview

    protected void btnCadLivro_Click(object sender, EventArgs e)
    {
        if (this.fuCapa.HasFile)
        {
            string filename = Path.GetFileName(this.fuCapa.PostedFile.FileName);
            string ext = Path.GetExtension(filename);
            if (ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".PNG" || ext == ".JPG" || ext == ".JPEG" || ext == ".gif" || ext == ".GIF")
            {
                try
                {
                    #region A

                    //FileStream fs = new FileStream(this.fuCapa.PostedFile.FileName, FileMode.Open, FileAccess.Read);
                    //byte[] capa = new byte[fs.Length];
                    //fs.Read(capa, 0, Convert.ToInt32(fs.Length));

                    //fs.Close();

                    //this.mLivro.NewBook(this.txtTituloLivro.Text, this.txtDescricaoLivro.Text, Convert.ToInt32(this.txtIdLivro.Text), Convert.ToInt32(this.ddlGeneroLivro.SelectedValue), Convert.ToInt32(this.ddlAutorLivro.SelectedValue), this.cbLivroLido.Checked, capa);

                    #endregion

                    #region B

                    Stream st = this.fuCapa.PostedFile.InputStream;
                    byte[] capa2 = new byte[st.Length];
                    st.Read(capa2, 0, Convert.ToInt32(st.Length));
                    this.mLivro.NewBook(this.txtTituloLivro.Text, this.txtDescricaoLivro.Text, Convert.ToInt32(this.txtIdLivro.Text), Convert.ToInt32(this.ddlGeneroLivro.SelectedValue), Convert.ToInt32(this.ddlAutorLivro.SelectedValue), this.cbLivroLido.Checked, capa2);

                    #endregion

                    #region C

                    /*Stream st = this.fuCapa.PostedFile.InputStream;
                    BinaryReader br = new BinaryReader(st);

                    this.mLivro.NewBook(this.txtTituloLivro.Text, this.txtDescricaoLivro.Text, Convert.ToInt32(this.txtIdLivro.Text), Convert.ToInt32(this.ddlGeneroLivro.SelectedValue), Convert.ToInt32(this.ddlAutorLivro.SelectedValue), this.cbLivroLido.Checked, br.ReadBytes((Int32)st.Length));
                    */
                    #endregion

                    this.CarregarLivros(Session[TIPO_CONSULTA_LIVROS] as bool?);
                }
                catch (BookException ex)
                {
                    ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('" + ex.Message + "');", true);
                }
                catch (Exception ex)
                {
                    ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Erro ao cadastrar livro');", true);
                }
            }
            else
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Arquivo inválido');", true);
            }
        }
        else
        {
            Response.Write("<script>alert('Selecione uma imagem');</script>");
        }
    }

    protected void gdvLivros_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            Image capa = (Image)e.Row.FindControl("imgCapa");
            Livro livro = e.Row.DataItem as Livro;

            if (livro.Capa != null)
            {
                MemoryStream ms = new MemoryStream(livro.Capa);
                string base64String = Convert.ToBase64String(livro.Capa, 0, livro.Capa.Length);
                capa.ImageUrl = "data:image/jpeg;base64," + base64String;
            }
        }
    }
    
    protected void gdvLivros_RowDeleting(object sender, GridViewDeleteEventArgs e)
    {
        int id;
        int.TryParse(this.gdvLivros.Rows[e.RowIndex].Cells[0].Text, out id);
        this.RemoverLivro(id);
        this.CarregarLivros(Session[TIPO_CONSULTA_LIVROS] as bool?);
    }
    
    protected void gdvLivros_RowCommand(object sender, GridViewCommandEventArgs e)
    {
        if (e.CommandName.Equals("EditarLivro"))
        {
            int row;
            int.TryParse(e.CommandArgument.ToString(), out row);
            this.txtIdLivro.Text = this.gdvLivros.Rows[row].Cells[0].Text;
            this.txtTituloLivro.Text = this.gdvLivros.Rows[row].Cells[1].Text;
            this.txtDescricaoLivro.Text = this.gdvLivros.Rows[row].Cells[2].Text;
            Label lido = (Label)gdvLivros.Rows[row].Cells[3].Controls[1];
            this.cbLivroLido.Checked = lido.Text.Equals("Sim") ? true : false;
            this.btnCadLivro.Visible = false;
            this.btnCancelarEdicao.Visible = true;
            this.btnAlterarLivro.Visible = true;

            this.pnlCadastroGenero.Visible = true;
        }
    }

    protected void btnCancelarEdicao_Click(object sender, EventArgs e)
    {
        this.LimparTela();
    }

    protected void btnAlterarLivro_Click(object sender, EventArgs e)
    {
        int id = int.Parse(this.txtIdLivro.Text);
        string titulo = this.txtTituloLivro.Text;
        int idAutor = Convert.ToInt32(this.ddlAutorLivro.SelectedValue);
        int idGenero = Convert.ToInt32(this.ddlGeneroLivro.SelectedValue);

        Stream st = this.fuCapa.PostedFile.InputStream;
        BinaryReader br = new BinaryReader(st);

        Livro livro = new Livro();
        livro.ISBN = id;
        livro.Titulo = titulo;
        livro.IdAutor = idAutor;
        livro.IdGenero = idGenero;
        livro.Lido = this.cbLivroLido.Checked;
        livro.Capa = br.ReadBytes((Int32)st.Length);

        this.mLivro.UpdateBook(livro);

        this.LimparTela();

        this.CarregarLivros(Session[TIPO_CONSULTA_LIVROS] as bool?);
    }

    protected void gdvLivros_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        this.gdvLivros.PageIndex = e.NewPageIndex;
        this.gdvLivros.DataBind();
    }

    #endregion
}