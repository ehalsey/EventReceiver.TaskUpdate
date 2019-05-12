<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="EventReceiver.TaskUpdateWeb.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Button ID="btnCreateEventReceiver" OnClick="btnCreateEventReceiver_Click" runat="server" Text="Create Event Receivers" />
        <asp:Button ID="btnRemoveEventReceiver" OnClick="btnRemoveEventReceiver_Click" runat="server" Text="Remove Event Receivers" />
        <asp:Button ID="btnRemoveEventReceiver1" OnClick="btnRemoveEventReceiver1_Click" runat="server" Text="Remove Event Receivers 1" Visible="false" />
        <asp:Button ID="btnRemoveEventReceiver2" OnClick="btnRemoveEventReceiver2_Click" runat="server" Text="Remove Event Receivers 2" Visible="false" />
        <asp:Button ID="btnRemoveEventReceiver3" OnClick="btnRemoveEventReceiver3_Click" runat="server" Text="Remove Event Receivers 3" Visible="false" />
    </div>
    </form>
</body>
</html>
