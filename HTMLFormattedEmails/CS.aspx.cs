using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using System.Data;

public partial class _Default : System.Web.UI.Page
{
    // protected void SendEmail(object sender, EventArgs e)
    // {
        // string body = this.PopulateBody("John",
            // "Fetch multiple values as Key Value pair in ASP.Net AJAX AutoCompleteExtender",
            // "http://www.aspsnippets.com/Articles/Fetch-multiple-values-as-Key-Value-pair-" + 
            // "in-ASP.Net-AJAX-AutoCompleteExtender.aspx",
            // "Here Mudassar Ahmed Khan has explained how to fetch multiple column values i.e." +
            // " ID and Text values in the ASP.Net AJAX Control Toolkit AutocompleteExtender"
            // + "and also how to fetch the select text and value server side on postback");
        // this.SendHtmlFormattedEmail("yogesh@inventive.in", "New article published!", body);
    // }

    // private string PopulateBody(string name, string email, string city, string description)
    // {
        // string body = string.Empty;
        // using (StreamReader reader = new StreamReader(Server.MapPath("~/mailer1.html")))
        // {
            // body = reader.ReadToEnd();
        // }
        // body = body.Replace("http://www.futureinstitutions.org/contactus/mailerContactus", "http://www.futureinstitutions.org/contactus/mailerContactus?name="+name+"&email="+email+"&city="+city+"");
        // //body = body.Replace("{Title}", title);
        // //body = body.Replace("{Url}", url);
        // //body = body.Replace("{Description}", description);
        // return body;
    // }

    // private void SendHtmlFormattedEmail(string recepientEmail, string sub, string body)
    // {

        // string userid = recepientEmail;
        // StringBuilder sb = new StringBuilder();

        // //sb.Append(@"<html xmlns='http://www.w3.org/1999/xhtml'><head>    <meta http-equiv='Content-Type' content='text/html; charset=UTF-8' />    <meta name='format-detection' content='telephone=no' />    <meta name='viewport' content='width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=no;' />    <meta http-equiv='X-UA-Compatible' content='IE=9; IE=8; IE=7; IE=EDGE' />    <link href='https://fonts.googleapis.com/css?family=Open+Sans:400,600' rel='stylesheet' />    <title>Registration Mail</title></head><body style='padding:0; margin:0; background:#fff'>    <table border='0' cellpadding='0' cellspacing='0' style='margin: 0; padding: 0; width:100%; font-family:Open Sans, Helvetica, Arial, sans-serif; background:#fff; color:#666; font-size:14px'>        <tr>            <td align='center' valign='top'>                <table class='content' border='0' cellpadding='0' cellspacing='0' style='width:100%; max-width:600px; background:#f9f9f9;'>                    <tr>                        <td>                            <img style='margin:0 auto;display:block;' src='http://www.futureinstitutions.org/images/logo.png' />                        </td>                    </tr>                    <tr>                                                <td align='center' style='padding:10px 15px'>                            <h1 style='margin:0; padding:0; font-weight:600; font-size:22px;color: #000;'>Future Institute of Management and Technology</h1>                        </td>                    </tr>                </table>                <table class='content' border='0' cellpadding='0' background='https://d3nn873nee648n.cloudfront.net/900x600/13623/220-ER486190.jpg' cellspacing='0' style='width:100%; background-image:url('https://d3nn873nee648n.cloudfront.net/900x600/13623/220-ER486190.jpg'); background-size:cover; background-position:center; background-repeat:no-repeat; max-width:600px'>                    <tr>                        <td style='width:50%; padding:30px 20px;background: #000;'>                            <h1 style='color:#fff; margin:0; padding:0; font-size:24px'>Register Now!</h1>                            <p style='margin:0; padding:20px 0; color:#fff; line-height:24px'>At Future we believe that learning must be stress-free. And that can only be made possible by creating a holistic environment where learning opportunities are created naturally.</p>                            <a href='http://www.futureinstitutions.org/contactus' target='_blank' style='display:inline-block; padding:10px 14px; font-weight:600; border-radius:0px; margin:10px 0; background:#3a72d3; color:#fff; text-decoration:none'>Click for Registration</a>                        </td>                        <td style='width:50%'></td>                    </tr>                </table>                <table class='content' border='0' cellpadding='0' cellspacing='0' style='width:100%; max-width:600px; background:#f9f9f9;'>                    <tr>                        <td style='padding:20px; line-height:24px; color:#333'>                            Thanks in advance                        </td>                    </tr>                </table>            </td>        </tr>    </table></body></html>");

        // const String FROM = "info@futureinstitutions.org";
        // const String FROMNAME = "Future Institutions";
        // string TO = userid;
        // const String SMTP_USERNAME = "yogesh@inventive.in";
        // const String SMTP_PASSWORD = "sNqGtVrx1yJ30Kdg";
        // const String CONFIGSET = "ConfigSet";
        // const String HOST = "smtp-relay.sendinblue.com";
        // const int PORT = 587;
        // string SUBJECT = "Hi "+sub;
        // string BODY = body;
        // MailMessage message = new MailMessage();
        // message.IsBodyHtml = true;
        // message.From = new MailAddress(FROM, FROMNAME);
        // message.To.Add(new MailAddress(TO));
        // //message.To.Add(new MailAddress('yogesh@inventiveinformatics.com'));
        // // message.To.Add(new MailAddress('yogesh@inventive.in'));
        // message.Subject = SUBJECT;
        // message.Body = BODY;
        // SmtpClient client =
            // new SmtpClient(HOST, PORT);
        // // Pass SMTP credentials
        // client.Credentials =
            // new NetworkCredential(SMTP_USERNAME, SMTP_PASSWORD);
        // // Enable SSL encryption
        // client.EnableSsl = true;

        // // Send the email. 
        // try
        // {
            // client.Send(message);
        // }
        // catch (Exception ex)
        // {
            // //Console.WriteLine('Error message: ' + ex.Message);
        // }
    // }

    protected void Upload(object sender, EventArgs e)
    {

        Random r = new Random();
        int num = r.Next();

        //Upload and save the file
        string excelPath = Server.MapPath("~/Files/") + Path.GetFileName(FileUpload1.PostedFile.FileName);
        FileUpload1.SaveAs(excelPath);

        string conString = string.Empty;
        string extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
        switch (extension)
        {
            case ".xls": //Excel 97-03
                conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                break;
            case ".xlsx": //Excel 07 or higher
                conString = ConfigurationManager.ConnectionStrings["Excel07+ConString"].ConnectionString;
                break;

        }
        conString = string.Format(conString, excelPath);
        using (OleDbConnection excel_con = new OleDbConnection(conString))
        {
            excel_con.Open();
            string sheet1 = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
            DataTable dtExcelData = new DataTable();

            //[OPTIONAL]: It is recommended as otherwise the data will be considered as String by default.
            dtExcelData.Columns.AddRange(new DataColumn[4] { new DataColumn("Name", typeof(string)),
                new DataColumn("Email", typeof(string)),
                new DataColumn("Mobile", typeof(string)),
                  new DataColumn("City", typeof(string))});

            using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", excel_con))
            {
                oda.Fill(dtExcelData);
            }
            excel_con.Close();

            string consString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
            using (SqlConnection con = new SqlConnection(consString))
            {
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                {
                    //Set the database table name
                    sqlBulkCopy.DestinationTableName = "dbo.tblPersons";

                    //[OPTIONAL]: Map the Excel columns with that of the database table
                    sqlBulkCopy.ColumnMappings.Add("Name", "Name");
                    sqlBulkCopy.ColumnMappings.Add("Email", "Email");
                    sqlBulkCopy.ColumnMappings.Add("Mobile", "Mobile");
                    sqlBulkCopy.ColumnMappings.Add("City", "City");
                    con.Open();
                    sqlBulkCopy.WriteToServer(dtExcelData);
                    con.Close();
                }
            }
        }
        UpdateTransNumber(num.ToString());
    }

    public void UpdateTransNumber(string num)
    {
        string consString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
        SqlDataAdapter dad = new SqlDataAdapter("update tblPersons set TransNumber='"+ num + "' where TransNumber is null;select * from tblPersons where TransNumber = '" + num + "' order by id;", consString);
        DataTable dt = new DataTable();
        dad.Fill(dt);
        if(dt.Rows.Count>0)
        {
            for(int i=0;i<dt.Rows.Count;i++)
            {
                string email = dt.Rows[i]["Email"].ToString();
                string name = dt.Rows[i]["Name"].ToString();
                string city = dt.Rows[i]["City"].ToString();
                string body = this.PopulateBody(name, email, city,"");
                this.SendHtmlFormattedEmail(email, name, body);
            }
        }
        dt.Clear();
        dad.Dispose();
    }
}