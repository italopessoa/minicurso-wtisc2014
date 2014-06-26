<%@ Page Title="ASP.NET WTISC 2014 - Autores" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="frmAutor.aspx.cs" Inherits="Default2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="content" ContentPlaceHolderID="cphContent" runat="Server">
    <asp:Panel ID="pnlCadastroAutor" runat="server">
        <asp:Table ID="tbCadastroAutor" runat="server">
            <asp:TableRow ID="tbrIdAutor" runat="server">
                <asp:TableCell runat="server">
                    <asp:Label ID="lblIdAutor" runat="server" Text="Id:"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server">
                    <asp:TextBox ID="txtIdAutor" runat="server" Width="50px" ReadOnly="true" Enabled="false"></asp:TextBox>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow ID="tbrNomeAutor" runat="server">
                <asp:TableCell runat="server">
                    <asp:Label ID="lblNomeAutor" runat="server" Text="Nome:"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server">
                    <asp:TextBox ID="txtNomeAutor" runat="server"></asp:TextBox>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow runat="server">
                <asp:TableCell runat="server" ColumnSpan="2">
                    <asp:Button ID="btnCadAutor" runat="server" Text="Cadastrar" ToolTip="Cadastrar novo Autor" OnClick="btnCadAutor_Click" />
                    <asp:Button ID="btnAlterarAutor" runat="server" Text="Alterar" ToolTip="Atualizar Autor" Visible="false" OnClick="btnAlterarAutor_Click" />
                    <asp:Button ID="btnCancelarEdicao" runat="server" Text="Cancelar" ToolTip="Cancelar edição" Visible="false" OnClick="btnCancelarEdicao_Click" />
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </asp:Panel>

    <asp:Panel ID="pnlGenerosCadastrados" runat="server">
        <asp:GridView ID="gdvAutores" runat="server" AutoGenerateColumns="false" OnRowDeleting="gdvAutores_RowDeleting"
            OnRowEditing="gdvAutores_RowEditing" OnRowDataBound="gdvAutores_RowDataBound" OnRowCommand="gdvAutores_RowCommand"
            OnPageIndexChanging="gdvAutores_PageIndexChanging" AllowPaging="true" PageSize="5" EmptyDataText="Nenhum autor foi cadastrado.">
            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#EFF3FB" />
            <PagerStyle BackColor="#bbbbbb" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#bbbbbb" Font-Bold="True" ForeColor="White" HorizontalAlign="Left" />
            <EditRowStyle BackColor="#2461BF" />
            <AlternatingRowStyle BackColor="White" />
            <EmptyDataRowStyle ForeColor="#ff0000" BackColor="Yellow" />
            <Columns>
                <asp:BoundField DataField="Id" HeaderText="Id" ItemStyle-Width="70px" ItemStyle-HorizontalAlign="Right" />
                <asp:BoundField DataField="Nome" HeaderText="Gênero" ItemStyle-Width="200px" />
                <asp:ButtonField ButtonType="Image" ImageUrl="~/imagens/1399616222_pen.png" CommandName="EditarAutor" />
                <asp:CommandField ButtonType="Image"
                    DeleteText="Remover" DeleteImageUrl="~/imagens/1399613611_Streamline-70.png" ShowDeleteButton="true" />
            </Columns>
        </asp:GridView>
    </asp:Panel>
</asp:Content>

