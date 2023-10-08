<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="FileUpload.aspx.cs" Inherits="BingoWeb.FileUpload" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>CSV Data</title>
    <link rel="stylesheet" type="text/css" href="~/Styles/styles.css" />
    <!-- Include your existing CSS file -->
    <link rel="stylesheet" type="text/css" href="styles.css" />
    <!-- Include the new CSS file -->
    <script type="text/javascript">
        function showWinnerMessage() {
            var overlay = document.getElementById('winnerOverlay');
            overlay.style.display = 'flex';

            // Add a delay of 5 seconds (5000 milliseconds)
            setTimeout(function () {
                var overlay = document.getElementById('winnerOverlay');
                overlay.style.display = 'none';

                // Show the gvOutput GridView after the delay
                var gvOutput = document.getElementById('gvOutput');
                gvOutput.style.display = 'grid-view';
            }, 2000); // Hide the message after 5 seconds (adjust as needed)
        }


    </script>

</head>
<body>
    <form id="form1" runat="server">
        <div class="container">
            <h1>Bingo!!</h1>



            <hr />

            <div class="btn-container">
                <div class="btn-container">
                    <asp:FileUpload ID="fileUpload" runat="server" />
                    <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click" CssClass="btn primary" /><br />
                    <p>
                        <asp:Label ID="lblError" runat="server" ForeColor="Red" />
                    </p>
                    <p>
                        <asp:Label ID="lblInfo" runat="server" ForeColor="Blue" />
                    </p>
                </div>
                <asp:TextBox ID="txtCount" runat="server" placeholder="Enter Count" />
                <asp:Button ID="btnSelectRandom" runat="server" Text="Generate" OnClick="btnSelectRandom_Click" CssClass="btn primary" /><br />
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
                <asp:Button ID="btnReset" runat="server" Text="Reset" OnClick="btnReset_Click" CssClass="btn secondary" />
                <asp:Button ID="btnPDF" runat="server" Text="Export to PDF" OnClick="btnPDF_Click" CssClass="btn secondary" />
                <asp:Button ID="btnExcel" runat="server" Text="Export to Excel" OnClick="btnExcel_Click" CssClass="btn secondary" />
                <asp:Button ID="btnEmail" runat="server" Text="Send PDF" Visible="false" OnClick="btnEmail_Click" CssClass="btn secondary" />
            </div>
        </div>
        <div id="winnerOverlay" class="overlay" style="display: none;">
            <div id="winnerMessage" class="winner-message">Here is your Winner!!!</div>
        </div>
    </form>

</body>
</html>
