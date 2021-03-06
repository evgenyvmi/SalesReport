﻿<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebApplication1.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">  
  
<html xmlns="http://www.w3.org/1999/xhtml">  
<head runat="server">  
    <title></title>  
</head>  
<body bgcolor="Silver">  
    <form id="form1" runat="server">  
    <div style="height: 263px">
        <h3>Choose the time period to create report from</h3>
        <p>from</p>  
        <asp:TextBox ID="TextBox1" runat="server" Columns="2" TextMode="Date" Rows="2"/>
        <p>to</p>  
        <asp:TextBox ID="TextBox2" runat="server" Columns="2" TextMode="Date" Rows="2"/>
        <asp:Button ID="Button1" runat="server"  Text="Choose range and create file" OnClick="Button1_Click"/>  
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
                <asp:BoundField DataField="OrderID" HeaderText="Order ID" />  
                <asp:BoundField DataField="OrderDate" HeaderText="Order Date" /> 
                <asp:BoundField DataField="Name" HeaderText="Name" />  
                <asp:BoundField DataField="Quantity" HeaderText="Quantity" />
                <asp:BoundField DataField="UnitPrice" HeaderText="Unit Price" /> 
            </Columns>  
        </asp:GridView>  
        <br />  
    </div>  
    </form>  
</body>  
</html>  