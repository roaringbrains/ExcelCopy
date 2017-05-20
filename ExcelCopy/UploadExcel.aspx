<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UploadExcel.aspx.cs" Inherits="ExcelCopy.UploadExcel" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server" enctype="multipart/form-data">
    <div>
    
        <asp:FileUpload ID="FileUpload1" runat="server" />
        <asp:Button ID="UploadButton1" runat="server" Text="Upload Report 1" OnClick="UploadButton1_Click" />
        <asp:Label ID="StatusLabel1" runat="server" Text=""></asp:Label>
        <br /><br />
    
        <asp:FileUpload ID="FileUpload2" runat="server" />
        <asp:Button ID="UploadButton2" runat="server" Text="Upload Report 2" OnClick="UploadButton2_Click" />
        <asp:Label ID="StatusLabel2" runat="server" Text=""></asp:Label>
        <br />
        <br />
    
        <asp:FileUpload ID="FileUpload3" runat="server" />
        <asp:Button ID="UploadButton3" runat="server" Text="Upload DevTools ATTRI" OnClick="UploadButton3_Click" />
        <asp:Label ID="StatusLabel3" runat="server" Text=""></asp:Label>
        <br />
        <br />
        <br />
        <br />
    
        <asp:Button ID="Calculate" runat="server" Text="Do The Magic\Compute" OnClick="Calculate_Click" />
        <br />
        <br /><br />
        <br />
        <asp:Label ID="FinalStatus" runat="server" Text=""></asp:Label>
        <br />
        <br />
        <br />
        <br />
        <br />
        <asp:GridView ID="GridView1" runat="server">
        </asp:GridView>
        <br /><br />

    </div>
    </form>
</body>
</html>
