<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="getDataFromSql.aspx.cs" Inherits="ExcelCopy.getDataFromSql" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:GridView ID="GridView1" runat="server"></asp:GridView>
        <br />
        <br />
        <br />
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Button" />
        <br />
    
    </div>
    </form>
</body>
</html>
