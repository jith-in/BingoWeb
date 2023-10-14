using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CsvHelper;
using CsvHelper.Configuration;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Excel = Microsoft.Office.Interop.Excel;
namespace BingoWeb
{
    public partial class FileUpload : System.Web.UI.Page
    {
        private List<Transaction> transactions;
        private string path;
        private string fullPath;
        private string strExcelPath;
        bool isAutoUpload = Convert.ToBoolean(ConfigurationManager.AppSettings["Autoupload"]?.ToString());
        bool isCorrespondentSpecfic = Convert.ToBoolean(ConfigurationManager.AppSettings["CorrespondentSpecfic"]?.ToString());
        int ifirstprice = Convert.ToInt32(ConfigurationManager.AppSettings["FirstPrizecount"].ToString());
        int isecondprice = Convert.ToInt32(ConfigurationManager.AppSettings["SecondPrizeCount"]?.ToString());
        string strFirstprizetext = ConfigurationManager.AppSettings["FirstPrizetxt"].ToString();
        string strSecondprizetext = ConfigurationManager.AppSettings["SecondPrizetxt"]?.ToString();
        string strThirdprizetext = ConfigurationManager.AppSettings["ThirdPrizetxt"]?.ToString();
        string strColourFirst = ConfigurationManager.AppSettings["ColourFirst"]?.ToString();
        string strColourSecond = ConfigurationManager.AppSettings["ColourSecond"]?.ToString();
        string csvFilePath = ConfigurationManager.AppSettings["CsvUploadPath"]?.ToString();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadData();
            }
            else
            {
                transactions = (List<Transaction>)Session["Transactions"];
            }
        }

        private void LoadData()
        {
            transactions = new List<Transaction>();
            Session["Transactions"] = transactions;

            // Initialize the list of selected transactions
            Session["SelectedTransactions"] = new List<Transaction>();

            if (isAutoUpload)
            {

                AutoUpload();
            }
            else
            {
                fileUpload.Visible = true;
                btnUpload.Visible = true;
                BindGridView();
            }
        }

        private void AutoUpload()
        {
            
            // Check if a specific CSV file path is configured
            if (!string.IsNullOrEmpty(csvFilePath) && File.Exists(csvFilePath))
            {
                try
                {
                    using (var reader = new StreamReader(csvFilePath))
                    using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
                    {
                        csv.Configuration.HeaderValidated = null;
                        csv.Configuration.MissingFieldFound = null;

                        // Read records
                        var records = csv.GetRecords<Transaction>().ToList();

                        // Handle null or empty values for the SpecialPrize property manually
                        foreach (var record in records)
                        {
                            //if (string.IsNullOrEmpty(record.RESULT))
                            //{
                            //    // Handle null or empty value here, e.g., set a default value
                            //    record.RESULT = "";
                            //}
                        }

                        transactions = records;

                        Session["Transactions"] = transactions;

                        // Clear the selection when a new file is uploaded
                        Session["SelectedTransactions"] = new List<Transaction>();
                        lblRecordCount.Visible = true;
                        lblRecordCount.Text = $"Uploaded Records: {transactions.Count}";

                        BindGridView();
                    }
                }
                catch (Exception ex)
                {
                    lblError.Text = $"Error: {ex.Message}";
                }
            }
            else
            {
                lblError.Text = "CSV file path " +csvFilePath + "  is not configured or the file does not exist.";
            }
        }


        private void BindGridView()
        {
            // Ensure that the GridView is bound to the correct data source
            var selectedTransactions = (List<Transaction>)Session["SelectedTransactions"];
            gvOutput.DataSource = selectedTransactions;
            gvOutput.DataBind();
            foreach (GridViewRow row in gvOutput.Rows)
            {
                string prize = row.Cells[6].Text;


                if (HttpUtility.HtmlDecode(prize?.Trim()).Equals(HttpUtility.HtmlDecode(strFirstprizetext), StringComparison.OrdinalIgnoreCase))
                {
                    row.Style["background-color"] = strColourFirst; // Set the background color to yellow
                    row.Style["font-weight"] = "bold"; // Set the font weight to bold
                }
                if(string.Equals(HttpUtility.HtmlDecode(prize?.Trim()), HttpUtility.HtmlDecode(strSecondprizetext?.Trim()), StringComparison.OrdinalIgnoreCase))
                {
                    row.Style["background-color"] = strColourSecond; // Set the background color to yellow
                    row.Style["font-weight"] = "bold"; // Set the font weight to bold
                }

            }
            // Bind the original gridView (gridView) separately
            gridView.DataSource = transactions;
            gridView.DataBind();
        }

        protected void btnSelectRandom_Click(object sender, EventArgs e)
        {
            // Your existing code for checking the count input
            if (int.TryParse(txtCount.Text, out int count))
            {
                Random random = new Random();
                var selectedTransactions = (List<Transaction>)Session["SelectedTransactions"];

                
                    string desiredCorrespondentsSetting = ConfigurationManager.AppSettings["DesiredCorrespondents"];
                    List<string> desiredCorrespondents = desiredCorrespondentsSetting.Split(',').ToList();


                    // Filter transactions by the desired "CORRESPONDENT" values
                    var transactionsToSelectFrom = transactions.Where(t => desiredCorrespondents.Contains(t.CORRESPONDENT)).ToList();
                
                while (count > 0)
                {
                    
                    var remainingTransactions = transactionsToSelectFrom
                        .Except(selectedTransactions)
                        .ToList();

                    if (!isCorrespondentSpecfic)
                    {
                        var transactions = GetTransactions(selectedTransactions);
                        remainingTransactions = transactions;
                    }

                    if (remainingTransactions.Count == 0)
                    {
                        lblError.Text = "Correspondent not Found or Incorrect configuration.";
                        break; // No more records to select
                    }

                    var newIndex = random.Next(remainingTransactions.Count);
                    var selectedTransaction = remainingTransactions[newIndex];

                    // Determine the prize for the selected transaction
                    string prize;
                    if (selectedTransactions.Count < ifirstprice)
                    {
                        prize = strFirstprizetext;
                        // Add code here to show overlay message
                        Page.ClientScript.RegisterStartupScript(this.GetType(), "ShowWinnerMessage",
                            "showWinnerMessage();", true);
                    }
                    else if (selectedTransactions.Count < isecondprice)
                    {
                        prize = strSecondprizetext;
                    }
                    else
                    {
                        prize = strThirdprizetext; // No specific prize for the rest
                    }

                    // Update the RESULT column with the prize
                    selectedTransaction.RESULT = prize;

                    selectedTransactions.Add(selectedTransaction);
                    count--;
                }

                // Remove the selected transactions from the original transactions list
                transactions.RemoveAll(t => selectedTransactions.Contains(t));

                // Update the session variable with the modified transactions list
                Session["Transactions"] = transactions;
                Session["SelectedTransactionsWithResult"] = selectedTransactions;
                // Refresh both GridViews
                BindGridView();
            }
            else
            {
                lblError.Text = "Please enter a valid count.";
            }
        }

        private List<Transaction>GetTransactions(List<Transaction> selectedTransactions)
        {
            List <Transaction> remainingTransactions = transactions
              .Except(selectedTransactions)
              .ToList();
            return remainingTransactions;
        }










        protected void gridView_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gridView.PageIndex = e.NewPageIndex;
            BindGridView(); // Rebind the gridView after changing the page index
        }

        protected void gvOutput_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gvOutput.PageIndex = e.NewPageIndex;
            BindGridView(); // Rebind the gvOutput after changing the page index
        }
        protected void btnReset_Click(object sender, EventArgs e)
        {
            // Clear uploaded data
            transactions = new List<Transaction>();
            Session["Transactions"] = transactions;

            // Clear selected transactions
            Session["SelectedTransactions"] = new List<Transaction>();
            Session["SelectedTransactionsWithResult"] = new List<Transaction>();
            // Clear error message
            lblError.Text = "";
            lblInfo.Text = "";
            // Clear record count label
            lblRecordCount.Text = "Uploaded Records: 0";

            // Reset GridView and gvOutput
            BindGridView();

            // Optionally, clear the file upload control by creating a new instance
            lblRecordCount.Visible = false;

            // Optionally, clear the count textbox
            txtCount.Text = "";
            if (isAutoUpload)
            {
                LoadData();
            }
        }




        public void btnUpload_Click(object sender, EventArgs e)
        {
            if (fileUpload.HasFile)
            {
                try
                {
                    using (var reader = new StreamReader(fileUpload.PostedFile.InputStream))
                    using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
                    {
                        csv.Configuration.HeaderValidated = null; //
                        csv.Configuration.MissingFieldFound = null; //

                        // Read records
                        var records = csv.GetRecords<Transaction>().ToList();

                        // Handle null or empty values for the SpecialPrize property manually
                        foreach (var record in records)
                        {
                            //if (string.IsNullOrEmpty(record.RESULT))
                            //{
                            //    // Handle null or empty value here, e.g., set a default value
                            //    record.RESULT = "";
                            //}
                        }

                        transactions = records;

                        Session["Transactions"] = transactions;

                        // Clear the selection when a new file is uploaded
                        Session["SelectedTransactions"] = new List<Transaction>();

                        // Refresh both GridViews
                        BindGridView();
                        lblRecordCount.Visible = true;
                        lblRecordCount.Text = $"Uploaded Records: {transactions.Count}";
                    }
                }
                catch (Exception ex)
                {
                    lblError.Text = $"Error: {ex.Message}";
                }
            }
            else
            {
                lblError.Text = "Please select a file to upload.";
            }
        }
        public void btnPDF_Click(object sender, EventArgs e)
        {
            var outputData = (List<Transaction>)Session["SelectedTransactionsWithResult"];
            if (outputData?.Count > 0)
            {
                if (ConfigurationManager.AppSettings.AllKeys.Contains("PDFDownloadPath"))
                {

                    var output = ConvertListToDataTable(outputData);
                    var op = ExportToPdf(output, false);

                }
                else
                {
                    var msg = "PDF Download path not configured. Please update path in appconfig";
                    string script = $"alert('{msg}');";
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", script, true);
                }

            }
            else
                lblError.Text = "No data available";
        }
        public DataTable ConvertListToDataTable(List<Transaction> transactions)
        {
            DataTable dt = new DataTable();

            // Define columns based on Transaction properties
            dt.Columns.Add("TXNDATE", typeof(string));
            dt.Columns.Add("REFNO", typeof(string));
            dt.Columns.Add("CUSTOMERNAME", typeof(string));
            dt.Columns.Add("IDNO", typeof(string));
            dt.Columns.Add("AMOUNT", typeof(string));
            dt.Columns.Add("CORRESPONDENT", typeof(string));
            dt.Columns.Add("RESULT", typeof(string));

            // Populate the DataTable with data from the list of Transaction objects
            foreach (var transaction in transactions)
            {
                dt.Rows.Add(
                    transaction.TXNDATE,
                    transaction.REFNO,
                    transaction.CUSTOMERNAME,
                    transaction.IDNO,
                    transaction.AMOUNT,
                    transaction.CORRESPONDENT,
                    transaction.RESULT
                );
            }

            return dt;
        }

        public string ExportToPdf(DataTable myDataTable, bool isMail)
        {
            DataTable dt = myDataTable;
            Document pdfDoc = new Document();
            Font font13 = FontFactory.GetFont("ARIAL", 13);
            Font font6 = FontFactory.GetFont("ARIAL", 6);
            Font headerFont = FontFactory.GetFont("HELVETICA", 15);
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            try
            {
                path = ConfigurationManager.AppSettings["PDFDownloadPath"].ToString();
                fullPath = Path.Combine(path, "TestExchange_Promotion_Draw_Results" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf");
                PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(fullPath, FileMode.Create));
                pdfDoc.Open();

                if (dt.Rows.Count > 0)
                {
                    PdfPTable PdfTable = new PdfPTable(1);

                    PdfPCell PdfPCell = new PdfPCell();
                    //string imageURL = @".\Sample File\DEX_2.png";

                    string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    string imageURL = Path.Combine(baseDirectory, "Sample File", "DEX_2.png");
                    BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arial.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    iTextSharp.text.Font font = new iTextSharp.text.Font(bf, 8);

                    iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imageURL);
                    jpg.ScaleToFit(pdfDoc.PageSize.Width, 75);
                    pdfDoc.Add(jpg);
                    string texttoDisplay = "Test Exchange Promotion Draw Results";
                    Paragraph para = new Paragraph(texttoDisplay, headerFont);
                    para.Alignment = Element.ALIGN_CENTER;
                    pdfDoc.Add(para);



                    PdfTable = new PdfPTable(dt.Columns.Count);
                    PdfTable.SpacingBefore = 25f;
                    for (int columns = 0; columns <= dt.Columns.Count - 1; columns++)
                    {
                        PdfPCell = new PdfPCell(new Phrase(new Chunk(dt.Columns[columns].ColumnName, font6)));
                        PdfTable.AddCell(PdfPCell);
                    }

                    for (int rows = 0; rows <= dt.Rows.Count - 1; rows++)
                    {
                        for (int column = 0; column <= dt.Columns.Count - 1; column++)
                        {

                            PdfPCell = new PdfPCell(new Phrase(new Chunk(Encoding.Unicode.GetString(Encoding.Unicode.GetBytes(dt.Rows[rows][column].ToString())), font)));
                            PdfTable.AddCell(PdfPCell);
                        }
                    }
                    pdfDoc.Add(PdfTable);
                }
                pdfDoc.AddCreationDate();
                pdfDoc.Close();
                if (!isMail)
                {
                    var msg = "Exported to " + fullPath.Replace("\\", "\\\\"); ;
                    string script = $"alert('{msg}');";
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", script, true);
                }
                return fullPath;



            }
            catch (DocumentException de)
            {
                lblError.Text = "Exception :" + de.Message.ToString();
                return "";
            }

            catch (IOException ioEx)
            {
                lblError.Text = "Exception :" + ioEx.Message.ToString();
                return "";
            }

            catch (Exception ex)
            {
                lblError.Text = "Exception :" + ex.Message.ToString();
                return "";
            }
        }

        public void btnExcel_Click(object sender, EventArgs e)
        {
            var outputData = (List<Transaction>)Session["SelectedTransactionsWithResult"];
            if (outputData?.Count > 0)
            {
                if (ConfigurationManager.AppSettings.AllKeys.Contains("ExcelDownloadPath"))
                {
                    string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");

                    string filepath = timestamp + "_" + "data.xlsx";

                    strExcelPath = ConfigurationManager.AppSettings["ExcelDownloadPath"].ToString() + filepath;
                    var output = ConvertListToDataTable(outputData);
                    ExportToExcel(output, strExcelPath);
                    var msg = "Exported to " + strExcelPath.Replace("\\", "\\\\"); ;
                    string script = $"alert('{msg}');";
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", script, true);

                }
                else
                {
                    var msg = "Excel Download path not configured. Please update path in config";
                    string script = $"alert('{msg}');";
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", script, true);

                }


            }
            else
                lblError.Text = "No data available";




        }
        public void ExportToExcel(DataTable dtExcel, string strExcelPath)
        {
            // Create a new Excel application
            Excel.Application excel = new Excel.Application();

            // Create a new workbook
            Excel.Workbook workbook = excel.Workbooks.Add();

            // Create a new worksheet
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            // Set the column headers
            for (int i = 0; i < dtExcel.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1] = dtExcel.Columns[i].ColumnName;
            }

            // Set the cell values
            for (int i = 0; i < dtExcel.Rows.Count; i++)
            {
                for (int j = 0; j < dtExcel.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dtExcel.Rows[i][j].ToString();
                }
            }

            workbook.SaveAs(strExcelPath);
            // Save the workbook


            // Close the workbook and release the resources
            workbook.Close();
            excel.Quit();




        }

        public void btnEmail_Click(object sender, EventArgs e)
        {
            SendEmailWithAttachment();
        }
        protected void SendEmailWithAttachment()
        {
            // Replace these values with your SMTP server and credentials
            string smtpServer = "smtp.gmail.com";
            int smtpPort = 587; // Port for SMTP server (e.g., 587 for Gmail)
            string smtpUsername = "";
            string smtpPassword = "";

            string fromEmail = "";
            string toEmail = "";
            string subject = "";
            string body = "";
            string pdfMailPath = "";
            // Create a MailMessage object
            MailMessage mail = new MailMessage(fromEmail, toEmail, subject, body);

            // Generate the PDF file
            var outputData = (List<Transaction>)Session["SelectedTransactionsWithResult"];
            if (outputData?.Count > 0)
            {
                if (ConfigurationManager.AppSettings.AllKeys.Contains("PDFDownloadPath"))
                {
                    var output = ConvertListToDataTable(outputData);
                    pdfMailPath = ExportToPdf(output, true);
                    if (!string.IsNullOrEmpty(pdfMailPath))
                    {
                        // Create an attachment
                        Attachment attachment = new Attachment(pdfMailPath);
                        mail.Attachments.Add(attachment);
                    }
                    else
                    {
                        lblError.Text = "Error generating PDF.";
                        return;
                    }
                }
                else
                {
                    lblError.Text = "PDF Download path not configured. Please update path in appconfig";
                    return;
                }
            }
            else
            {
                lblError.Text = "No data available";
                return;
            }

            // Create a SmtpClient to send the email
            SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort);
            smtpClient.Credentials = new System.Net.NetworkCredential(smtpUsername, smtpPassword);

            // Enable SSL if your SMTP server requires it
            smtpClient.EnableSsl = true;

            try
            {
                // Send the email
                smtpClient.Send(mail);
                lblInfo.Text = "Email sent successfully!";
            }
            catch (Exception ex)
            {
                lblError.Text = "Error sending email: " + ex.Message;
                // Assuming you have a label control named lblErrorMessage in your ASP.NET markup
                // Assuming you have a label control named lblErrorMessage in your ASP.NET markup
                lblInfo.Text = @"
                    <span style='color: red; font-weight: bold;'>Provide Valid Credentials:</span> Make sure you are providing valid username and password credentials to authenticate with the SMTP server. The smtpUsername and smtpPassword variables in your code should contain the correct login credentials for your email account.

                    <span style='color: red; font-weight: bold;'>Check for 2-Step Verification:</span> If you are using Gmail or a similar service, and you have two-step verification enabled for your account, you may need to generate an 'App Password' specifically for your application. Regular email account passwords may not work in this case.

                    <span style='color: red; font-weight: bold;'>Network and Firewall:</span> Ensure that your network connection is stable and not blocked by a firewall or security software. Sometimes, network issues or firewall rules can prevent your application from connecting to the SMTP server.";


            }
            finally
            {
                // Clean up resources
                mail.Dispose();
                smtpClient.Dispose();
                if (File.Exists(pdfMailPath))
                {
                    try
                    {
                        // Delete the PDF file
                        File.Delete(pdfMailPath);

                    }
                    catch (Exception ex)
                    {
                        lblError.Text = "Error deleting PDF file: " + ex.Message;
                    }
                }
            }
        }



    }
}
