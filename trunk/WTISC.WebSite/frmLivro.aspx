<%@ Page Title="ASP.NET WTISC 2014 - Livros" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="frmLivro.aspx.cs" Inherits="frmLivro" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="content" ContentPlaceHolderID="cphContent" runat="Server">
    <asp:Panel ID="pnlCadastroGenero" runat="server">
        <asp:Table ID="tbCadastroGenero" runat="server">
            <asp:TableRow ID="tbrIdLivro" runat="server">
                <asp:TableCell runat="server">
                    <asp:Label ID="lblIdLivro" runat="server" Text="ISBN:"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server">
                    <asp:TextBox ID="txtIdLivro" runat="server" Width="120px"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server">
                    <asp:Label ID="lblGeneroLivro" runat="server" Text="Gênero:"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server">
                    <asp:DropDownList ID="ddlGeneroLivro" runat="server"></asp:DropDownList>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow ID="tbrTituloLivro" runat="server">
                <asp:TableCell runat="server">
                    <asp:Label ID="lblTituloLivro" runat="server" Text="Título:"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server">
                    <asp:TextBox ID="txtTituloLivro" runat="server"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server">
                    <asp:Label ID="lblAutorLivro" runat="server" Text="Autor:"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server">
                    <asp:DropDownList ID="ddlAutorLivro" runat="server"></asp:DropDownList>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell runat="server">
                    <asp:Label ID="lblDescricaoLivro" runat="server" Text="Resumo:"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" ColumnSpan="4">
                    <asp:TextBox ID="txtDescricaoLivro" runat="server" TextMode="MultiLine" Width="300px" Height="80px" MaxLength="300"></asp:TextBox>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell runat="server">
                    <asp:Label runat="server" Text="Capa"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server">
                    <asp:FileUpload ID="fuCapa" runat="server" />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:CheckBox ID="cbLivroLido" runat="server" Text="Lido" />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow runat="server">
                <asp:TableCell runat="server" ColumnSpan="2">
                    <asp:Button ID="btnCadLivro" runat="server" Text="Cadastrar" ToolTip="Cadastrar novo gênero" OnClick="btnCadLivro_Click" />
                    <asp:Button ID="btnAlterarLivro" runat="server" Text="Alterar" ToolTip="Atualizar gênero" OnClick="btnAlterarLivro_Click" Visible="false" />
                    <asp:Button ID="btnCancelarEdicao" runat="server" Text="Cancelar" ToolTip="Cancelar edição" OnClick="btnCancelarEdicao_Click" Visible="false" />
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </asp:Panel>
    <asp:Panel ID="pnlLivrosCadastrados" runat="server">
        <asp:GridView ID="gdvLivros" runat="server" AutoGenerateColumns="false" Width="780px"
            AllowPaging="true" PageSize="2" EmptyDataText="Nenhum livro foi cadastrado." OnRowDataBound="gdvLivros_RowDataBound" OnRowCommand="gdvLivros_RowCommand" OnRowDeleting="gdvLivros_RowDeleting" OnPageIndexChanging="gdvLivros_PageIndexChanging">
            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#EFF3FB" />
            <PagerStyle BackColor="#bbbbbb" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#bbbbbb" Font-Bold="True" ForeColor="White" HorizontalAlign="Left" />
            <EditRowStyle BackColor="#2461BF" />
            <AlternatingRowStyle BackColor="White" />
            <EmptyDataRowStyle ForeColor="#ff0000" BackColor="Yellow" />
            <Columns>
                <asp:BoundField DataField="ISBN" HeaderText="ISBN" HeaderStyle-HorizontalAlign="Center" ItemStyle-Width="70px" ItemStyle-HorizontalAlign="Right" />
                <asp:BoundField DataField="Titulo" HeaderText="Título" HeaderStyle-HorizontalAlign="Center" ItemStyle-Width="200px" />
                <asp:BoundField DataField="Resumo" HeaderText="Resumo" HeaderStyle-HorizontalAlign="Center" ItemStyle-Width="450px" />
                <asp:TemplateField HeaderText="Lido" HeaderStyle-HorizontalAlign="Center" HeaderStyle-Width="90px" ItemStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <asp:Label ID="lblLivroLidoValue" runat="server" Text='<%# Convert.ToBoolean(Eval("Lido")) ? "Sim" : "Não" %>' />
                    </ItemTemplate>
                </asp:TemplateField>

                <asp:BoundField DataField="Autor.Nome" HeaderText="Autor" HeaderStyle-HorizontalAlign="Center" HeaderStyle-Width="200px" />
                <asp:BoundField DataField="Genero.Nome" HeaderText="Gênero" HeaderStyle-HorizontalAlign="Center" ItemStyle-Width="200px" />
                <asp:TemplateField HeaderText="Capa" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <asp:Image ID="imgCapa" runat="server" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:ButtonField ButtonType="Image" ImageUrl="~/imagens/1399616222_pen.png" CommandName="EditarLivro" />
                <asp:CommandField ButtonType="Image"
                    DeleteText="Remover" DeleteImageUrl="~/imagens/1399613611_Streamline-70.png" ShowDeleteButton="true" />
            </Columns>
        </asp:GridView>
    </asp:Panel>
</asp:Content>

