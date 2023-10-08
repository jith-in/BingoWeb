<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="FileUpload.aspx.cs" Inherits="BingoWeb.FileUpload" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>CSV Data</title>
    <link rel="stylesheet" type="text/css" href="~/Styles/styles.css" />
</head>
<body>
    <form id="form1" runat="server">
        <div class="container">
            <h1>Bingo!!</h1>

            <div class="btn-container">
                <asp:FileUpload ID="fileUpload" runat="server" />
                <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click" CssClass="btn" /><br />
                <p>
                    <asp:Label ID="lblError" runat="server" ForeColor="Red" />
                </p>
            </div>

            <hr />

            <div class="btn-container">
                <asp:TextBox ID="txtCount" runat="server" placeholder="Enter Count" />
                <asp:Button ID="btnSelectRandom" runat="server" Text="Generate" OnClick="btnSelectRandom_Click" CssClass="btn" /><br />
                <asp:Label ID="lblRecordCount" Visible="False" runat="server" Text="Uploaded Records: 0" />
            </div>

            <div class="grid-view">
                <asp:GridView ID="gridView" runat="server" OnPageIndexChanging="gridView_PageIndexChanging" AutoGenerateColumns="false" AllowPaging="true" PageSize="10" CssClass="grid-view">
                    <Columns>
                        <asp:BoundField DataField="TXNDATE" HeaderText="TXNDATE" />
                        <asp:BoundField DataField="REFNO" HeaderText="REFNO" />
                        <asp:BoundField DataField="CUSTOMERNAME" HeaderText="CUSTOMERNAME" />
                        <asp:BoundField DataField="IDNO" HeaderText="IDNO" />
                        <asp:BoundField DataField="AMOUNT" HeaderText="AMOUNT" />
                        <asp:BoundField DataField="CORRESPONDENT" HeaderText="CORRESPONDENT" />
                        <asp:BoundField DataField="RESULT" HeaderText="RESULT" Visible="false" />
                    </Columns>
                </asp:GridView>
                <asp:GridView ID="gvOutput" runat="server" OnPageIndexChanging="gvOutput_PageIndexChanging" AllowPaging="true" PageSize="6" CssClass="grid-view">
                </asp:GridView>
            </div>

            <div class="btn-container">
                <asp:Button ID="btnReset" runat="server" Text="Reset" OnClick="btnReset_Click" CssClass="btn"  />
                <asp:Button ID="btnPDF" runat="server" Text="Export to PDF" OnClick="btnPDF_Click" CssClass="btn" />
                <asp:Button ID="btnExcel" runat="server" Text="Export to Excel" OnClick="btnExcel_Click" CssClass="btn" />
            </div>
        </div>
    </form>
</body>
</html>
