<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Read.aspx.cs" Inherits="ExcelCopy.Read" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Start Reading" /><br />
        <asp:Label ID="Label1" runat="server" Text="Excel Location :D:\Test1.xlsx"></asp:Label></div>
        <br /><br /><br /><br /><br /><br />

        <asp:GridView ID="FileData" runat="server"></asp:GridView>

        <br />
        <br />
        <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Text="Export To Excel" />

    </form>
</body>
</html>
