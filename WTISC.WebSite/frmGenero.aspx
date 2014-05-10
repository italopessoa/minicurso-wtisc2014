<%@ Page Title="ASP.NET WTISC 2014 - Gêneros" Language="C#" AutoEventWireup="true" CodeFile="frmGenero.aspx.cs" Inherits="frmGenero" MasterPageFile="~/MasterPage.master" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="content" ContentPlaceHolderID="cphContent" runat="Server">
    <asp:Panel ID="pnlCadastroGenero" runat="server">
        <asp:Table ID="tbCadastroGenero" runat="server">
            <asp:TableRow ID="tbrIdGenero" runat="server">
                <asp:TableCell runat="server">
                    <asp:Label ID="lblIdGenero" runat="server" Text="Id:"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server">
                    <asp:TextBox ID="txtIdGenero" runat="server" Width="50px" ReadOnly="true" Enabled="false"></asp:TextBox>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow ID="tbrNomeGenero" runat="server">
                <asp:TableCell runat="server">
                    <asp:Label ID="lblNomeGenero" runat="server" Text="Nome:"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server">
                    <asp:TextBox ID="txtNomeGenero" runat="server"></asp:TextBox>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell runat="server">
                    <asp:Label ID="lblDescricaoGenero" runat="server" Text="Descrição:"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server">
                    <asp:TextBox ID="txtDescricaoGenero" runat="server" TextMode="MultiLine" MaxLength="300"></asp:TextBox>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow runat="server">
                <asp:TableCell runat="server" ColumnSpan="2">
                    <asp:Button ID="btnCadGenero" runat="server" Text="Cadastrar" ToolTip="Cadastrar novo gênero" OnClick="btnCadGenero_Click" />
                    <asp:Button ID="btnAlterarGenero" runat="server" Text="Alterar" ToolTip="Atualizar gênero" Visible="false" OnClick="btnAlterarGenero_Click" />
                    <asp:Button ID="btnCancelarEdicao" runat="server" Text="Cancelar" ToolTip="Cancelar edição" Visible="false" OnClick="btnCancelarEdicao_Click" />
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </asp:Panel>

    <asp:Panel ID="pnlGenerosCadastrados" runat="server">
        <asp:GridView ID="gdvGeneros" runat="server" AutoGenerateColumns="false" OnRowDeleting="gdvGeneros_RowDeleting" Width="600px"
            OnRowEditing="gdvGeneros_RowEditing" OnRowDataBound="gdvGeneros_RowDataBound" OnRowCommand="gdvGeneros_RowCommand"
            OnPageIndexChanging="gdvGeneros_PageIndexChanging" AllowPaging="true" PageSize="4" EmptyDataText="Nenhum gênero foi cadastrado.">
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
                <asp:BoundField DataField="Descricao" HeaderText="Descrição" ItemStyle-Width="450px" />
                <asp:ButtonField ButtonType="Image" ImageUrl="~/imagens/1399616222_pen.png" CommandName="EditarGenero" />
                <asp:CommandField ButtonType="Image"
                    DeleteText="Remover" DeleteImageUrl="~/imagens/1399613611_Streamline-70.png" ShowDeleteButton="true" />
            </Columns>
        </asp:GridView>
    </asp:Panel>
</asp:Content>
