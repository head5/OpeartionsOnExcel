<%@ Page Title="Home Page" Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebRole1._Default" %>


<form id="form1" runat="server">
    <asp:FileUpload ID="browseExcel" runat="server" Width="623px" />
&nbsp;&nbsp;
    <br />
    <br />
    <asp:DropDownList ID="DropDownList1" runat="server" Height="16px" Width="217px">
    </asp:DropDownList>
&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:Button ID="butShowDetails" runat="server" Font-Bold="True" Font-Italic="True" Font-Overline="False" OnClick="butShowDetails_Click" Text="Show Details" Width="168px" />
    <p>
        <asp:Table ID="tblResult" runat="server"  BackColor="#CCCCCC" BorderColor="#336699" BorderStyle="Solid" CellPadding="5" CellSpacing="8" BorderWidth="2px">
        </asp:Table>
    </p>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
</form>



