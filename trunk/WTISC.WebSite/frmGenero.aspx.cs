using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using WTISC._2014.Data;
using WTISC._2014.Data.Exceptions;
using WTISC._2014.Data.Management;

public partial class frmGenero : System.Web.UI.Page
{
    #region membros privados

    private MGenero mGenero = new MGenero();
    private const string SESSION_GENERO = "session_genero";

    #endregion

    protected void Page_Load(object sender, EventArgs e)
    {
        this.ExibirGenerosCadastrados();
    }

    #region botoes

    protected void btnCadGenero_Click(object sender, EventArgs e)
    {
        if (!String.IsNullOrEmpty(this.txtNomeGenero.Text))
        {
            if (!String.IsNullOrEmpty(this.txtDescricaoGenero.Text))
            {
                try
                {
                    this.mGenero.NewGender(this.txtNomeGenero.Text, this.txtDescricaoGenero.Text);
                    ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('"+String.Format("Gênero: \"{0}\" cadastrado com sucesso!",this.txtNomeGenero.Text)+"');", true);
                    this.ExibirGenerosCadastrados();
                }
                catch (GenderException ex)
                {
                    ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('"+ex.Message+"');", true);
                }

            }
            else
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Informe a descrição!');", true);
            }
        }
        else
        {
            ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Informe o nome!');", true);
        }
    }
    
    protected void btnAlterarGenero_Click(object sender, EventArgs e)
    {
        int id = int.Parse(this.txtIdGenero.Text);
        string nome = this.txtNomeGenero.Text;
        string descricao = this.txtDescricaoGenero.Text;

        Genero genero = new Genero() { Id = id, Nome = nome, Descricao = descricao };
        this.mGenero.Update(genero);

        this.btnAlterarGenero.Visible = false;
        this.btnCadGenero.Visible = true;
        this.btnCancelarEdicao.Visible = false;

        this.ExibirGenerosCadastrados();
    }

    protected void btnCancelarEdicao_Click(object sender, EventArgs e)
    {
        this.btnAlterarGenero.Visible = false;
        this.btnCadGenero.Visible = true;
        this.btnCancelarEdicao.Visible = false;

        this.txtIdGenero.Text = String.Empty;
        this.txtNomeGenero.Text = string.Empty;
        this.txtDescricaoGenero.Text = string.Empty;
    }

    #endregion

    #region metodos privados

    private void ExibirGenerosCadastrados()
    {
        Session[SESSION_GENERO] = this.mGenero.FindAll();
        this.gdvGeneros.DataSource = Session[SESSION_GENERO] as List<Genero>;
        this.gdvGeneros.DataBind();
    }
    
    private void RemoverGenero(int idGenero)
    {
        this.mGenero.Delete(idGenero);
    }

    #endregion

    #region gridview

    protected void gdvGeneros_RowDeleting(object sender, GridViewDeleteEventArgs e)
    {
        int id;
        int.TryParse(this.gdvGeneros.Rows[e.RowIndex].Cells[0].Text, out id);
        this.RemoverGenero(id);
        this.ExibirGenerosCadastrados();
    }
    
    protected void gdvGeneros_RowEditing(object sender, GridViewEditEventArgs e)
    {
        

        
    }

    protected void gdvGeneros_RowDataBound(object sender, GridViewRowEventArgs e)
    {
    }
    
    protected void gdvGeneros_RowCommand(object sender, GridViewCommandEventArgs e)
    {
        if (e.CommandName.Equals("EditarGenero"))
        {
            int row;
            int.TryParse(e.CommandArgument.ToString(), out row);
            this.txtIdGenero.Text = this.gdvGeneros.Rows[row].Cells[0].Text;
            this.txtNomeGenero.Text = this.gdvGeneros.Rows[row].Cells[1].Text;
            this.txtDescricaoGenero.Text = this.gdvGeneros.Rows[row].Cells[2].Text;

            this.btnCadGenero.Visible = false;
            this.btnCancelarEdicao.Visible = true;
            this.btnAlterarGenero.Visible = true;
        }
    }
    
    protected void gdvGeneros_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        this.gdvGeneros.PageIndex = e.NewPageIndex;
        this.gdvGeneros.DataBind();
    }

    #endregion
}