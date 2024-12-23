<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="COMInteropWebExample.Default" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>COM Interop Web Example</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <h1>COM Interop Example</h1>
            <asp:Button ID="btnCreateFile" runat="server" Text="Create File" OnClick="btnCreateFile_Click" />
            <br /><br />
            <asp:Label ID="lblStatus" runat="server" Text=""></asp:Label>
        </div>
    </form>
</body>
</html>
