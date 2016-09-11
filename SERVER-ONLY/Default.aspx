<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="SERVER_ONLY._Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:ScriptManager ID="ScriptManager1" runat="server">
            </asp:ScriptManager>
            <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                <ContentTemplate>
                    <div style="align-content:center;text-align:center;vertical-align:middle">
                        
                        <asp:Label ID="Label1" runat="server" Text="Time is Now: " Font-Size="Large"></asp:Label>
                        <asp:Label ID="nt" runat="server" Font-Size="Large" Font-Italic="True" ForeColor="#FF0066"></asp:Label>
                        <br />
                        <asp:Label ID="ERR" runat="server" ForeColor="Blue"></asp:Label>
                        <br />
                        <asp:Label ID="CTSTA" runat="server"></asp:Label>
                    </div>
                    <asp:Timer ID="Timer1" runat="server" Interval="1000">
                    </asp:Timer>
                </ContentTemplate>
            </asp:UpdatePanel>
        </div>
    </form>
</body>
</html>
