<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="errgrid.aspx.vb" Inherits="SERVER_ONLY.errgrid" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:GridView ID="GridView1" HeaderStyle-BackColor="#3AC0F2" HeaderStyle-ForeColor="White"
    runat="server" AutoGenerateColumns="false">
    <Columns>
        <asp:BoundField DataField="EID" HeaderText="Id" ItemStyle-Width="100" />
        <asp:BoundField DataField="ETIME" HeaderText="Name" ItemStyle-Width="200" />
        <asp:BoundField DataField="ERR" HeaderText="Country" ItemStyle-Width="500" />
    </Columns>
</asp:GridView>
    </div>
    </form>
</body>
</html>
