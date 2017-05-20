<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Form2.aspx.cs" Inherits="ExcelCopy.Form2" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
     <form id="form1" runat="server">
        <div>
            <div style="float: left;">
                <asp:FileUpload ID="FileUploadControl" runat="server" />
        <asp:Button ID="btnUpload" runat="server" 

        Text="Upload" OnClick="btnUpload_Click" />            
        </div>
            <div style="float: left;">
                <asp:Label ID="Label" runat="server" 

                Text="Is Header Exists?"></asp:Label>
                <asp:DropDownList ID="ddlIsHeaderExists" runat="server">
                    <asp:ListItem Value="Yes">Yes</asp:ListItem>
                    <asp:ListItem Value="No">No</asp:ListItem>
                </asp:DropDownList>
                <label runat="server" style="color: red;" 

                id="lblErrorMessage" visible="false"></label>
            </div>
            <div style="clear: both;padding-top:20px;"></div>
            <div>
                <asp:GridView ID="ExcelGridView" 

                runat="server" AllowPaging="false">
                </asp:GridView>
            </div>
        </div>
    </form>
</body>
</html>
