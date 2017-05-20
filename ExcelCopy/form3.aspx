<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="form3.aspx.cs" Inherits="ExcelCopy.form3" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="frm" runat="server">
    <div>
      <asp:Button ID="btnImport" Text="Import" OnClick="btnImport_Click" runat="server"/>
    </div>
    <br />
    <div>
        <asp:Label ID="lblMsg" runat="server" Font-Bold="true"></asp:Label>
        <br />
        <asp:GridView ID="gvImport" runat="server" AutoGenerateColumns="false">
            <EmptyDataTemplate>
                <div style="padding:10px">
                    Data not found!
                </div>
            </EmptyDataTemplate>
        </asp:GridView>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataKeyNames="RentalID" DataSourceID="SqlDataSource1">
            <Columns>
                <asp:BoundField DataField="RentalID" HeaderText="RentalID" InsertVisible="False" ReadOnly="True" SortExpression="RentalID" />
                <asp:BoundField DataField="MovieID" HeaderText="MovieID" SortExpression="MovieID" />
                <asp:BoundField DataField="CustomerID" HeaderText="CustomerID" SortExpression="CustomerID" />
                <asp:BoundField DataField="DateRented" HeaderText="DateRented" SortExpression="DateRented" />
            </Columns>
        </asp:GridView>
        <asp:EntityDataSource ID="EntityDataSource1" runat="server">
        </asp:EntityDataSource>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:MoviesDBConnectionString %>" SelectCommand="SELECT * FROM [Rental]"></asp:SqlDataSource>
    </div>
    </form>
</body>
</html>
