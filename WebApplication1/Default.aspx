﻿<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebApplication1.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">  
  
<html xmlns="http://www.w3.org/1999/xhtml">  
<head runat="server">  
    <title></title>  
</head>  
<body bgcolor="Silver">  
    <form id="form1" runat="server">  
    <div>  
        <asp:GridView ID="GridView1" AutoGenerateColumns="false" runat="server" CellPadding="6"  
            ForeColor="#333333" GridLines="None">  
            <AlternatingRowStyle BackColor="White" />  
            <EditRowStyle BackColor="#7C6F57" />  
            <FooterStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />  
            <HeaderStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />  
            <PagerStyle BackColor="#666666" ForeColor="White" HorizontalAlign="Center" />  
            <RowStyle BackColor="#E3EAEB" />  
            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />  
            <SortedAscendingCellStyle BackColor="#F8FAFA" />  
            <SortedAscendingHeaderStyle BackColor="#246B61" />  
            <SortedDescendingCellStyle BackColor="#D4DFE1" />  
            <SortedDescendingHeaderStyle BackColor="#15524A" />  
            <Columns>  
                <asp:BoundField DataField="id" HeaderText="id" />  
                <asp:BoundField DataField="Name" HeaderText="Name" />  
                <asp:BoundField DataField="City" HeaderText="City" />  
                <asp:BoundField DataField="Address" HeaderText="Address" />  
                <asp:BoundField DataField="Designation" HeaderText="Designation" />  
            </Columns>  
        </asp:GridView>  
        <br />  
                   <asp:Button ID="Button1" runat="server"  
            Text="Create Sales Report File" onclick="Button1_Click"/>  
    </div>  
    </form>  
</body>  
</html>  