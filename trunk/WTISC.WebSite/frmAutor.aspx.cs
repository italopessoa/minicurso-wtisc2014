using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using WTISC._2014.Data.Exceptions;
using WTISC._2014.Data.Management;
using WTISC._2014.Data;

public partial class Default2 : System.Web.UI.Page
{
    #region Membros privados

    private MAutor mAutor = new MAutor();
    private const string SESSION_AUTOR = "session_autor";

    #endregion

    protected void Page_Load(object sender, EventArgs e)
    {
        this.ExibirAutoresCadastrados();
    }

    #region botoes

    protected void btnCadAutor_Click(object sender, EventArgs e)
    {
        if (!String.IsNullOrEmpty(this.txtNomeAutor.Text))
        {
            try
            {
                this.mAutor.NewAuthor(this.txtNomeAutor.Text);
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('" + String.Format("Autor: \"{0}\" cadastrado com sucesso!", this.txtNomeAutor.Text) + "');", true);
                this.ExibirAutoresCadastrados();
            }
            catch (AuthorException ex)
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('" + ex.Message + "');", true);
            }
        }
        else
        {
            ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Informe o nome!');", true);
        }
    }

    protected void btnAlterarAutor_Click(object sender, EventArgs e)
    {
        int id = int.Parse(this.txtIdAutor.Text);
        string nome = this.txtNomeAutor.Text;

        Autor autor = new Autor() { Id = id, Nome = nome };
        this.mAutor.Update(autor);

        this.btnAlterarAutor.Visible = false;
        this.btnCadAutor.Visible = true;
        this.btnCancelarEdicao.Visible = false;

        this.ExibirAutoresCadastrados();
    }

    #endregion

    #region metodos privados

    private void ExibirAutoresCadastrados()
    {
        Session[SESSION_AUTOR] = this.mAutor.FindAll();
        this.gdvAutores.DataSource = Session[SESSION_AUTOR] as List<Autor>;
        this.gdvAutores.DataBind();
    }

    private void RemoverAutor(int idAutor)
    {
        this.mAutor.Delete(idAutor);
    }

    #endregion

    #region gridview

    protected void gdvAutores_RowDeleting(object sender, GridViewDeleteEventArgs e)
    {
        int id;
        int.TryParse(this.gdvAutores.Rows[e.RowIndex].Cells[0].Text, out id);
        this.RemoverAutor(id);
        this.ExibirAutoresCadastrados();
    }

    protected void gdvAutores_RowEditing(object sender, GridViewEditEventArgs e)
    {



    }

    protected void btnCancelarEdicao_Click(object sender, EventArgs e)
    {
        this.btnAlterarAutor.Visible = false;
        this.btnCadAutor.Visible = true;
        this.btnCancelarEdicao.Visible = false;

        this.txtIdAutor.Text = String.Empty;
        this.txtNomeAutor.Text = string.Empty;
    }

    protected void gdvAutores_RowDataBound(object sender, GridViewRowEventArgs e)
    {
    }

    protected void gdvAutores_RowCommand(object sender, GridViewCommandEventArgs e)
    {
        if (e.CommandName.Equals("EditarAutor"))
        {
            int row;
            int.TryParse(e.CommandArgument.ToString(), out row);
            this.txtIdAutor.Text = this.gdvAutores.Rows[row].Cells[0].Text;
            this.txtNomeAutor.Text = this.gdvAutores.Rows[row].Cells[1].Text;

            this.btnCadAutor.Visible = false;
            this.btnCancelarEdicao.Visible = true;
            this.btnAlterarAutor.Visible = true;
        }
    }

    protected void gdvAutores_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        this.gdvAutores.PageIndex = e.NewPageIndex;
        this.gdvAutores.DataBind();
    }

    #endregion
}