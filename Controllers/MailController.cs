using System;
using System.Configuration;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Mvc;
using Gemplastics.Views.Mail;

namespace Gemplastics.Controllers
{
    public class MailController : Controller
    {
        public ActionResult QuoteQuery()
        {
            var model = new QuoteQueryModel();
            
            return View(model);
        }

        public ActionResult Send(QuoteQueryModel model)
        {
            if (string.IsNullOrEmpty(model.Person) ||
               string.IsNullOrEmpty(model.Tel) ||
               string.IsNullOrEmpty(model.Email) ||
               IsValidEmail(model.Email) == false ||
               string.IsNullOrEmpty(model.Person) ||
               string.IsNullOrEmpty(model.Query))
            {
                model.Error = "Please complete the mandatory details!";
                return View("QuoteQuery", model);
            }

            model.Error = string.Empty;
            
            SendRequest(
                GemplasticsMail(model.Person, model.Tel, model.Email, model.Query, model.Width, model.Length, model.Microns, model.PlasticColor, model.Color1, model.Color2, model.Color3),
                ConfigurationManager.AppSettings["EmailTo"], model.Email, "Query from Website!");

            SendRequest(
                "Thank You for visiting Gemplastics! We will contact you shortly!" + @"<br /><br />" + "Gemplastics",
                model.Email, ConfigurationManager.AppSettings["EmailTo"], "E-mail from Gemplastic's website!");
            
            return View("Sent");
        }

        public bool IsValidEmail(string emlAddress)
        {
            //create Regular Expression Match pattern object
            Regex myRegex = new Regex("^([a-zA-Z0-9_\\-\\.]+)@[a-z0-9-]+(\\.[a-z0-9-]+)*(\\.[a-z]{2,3})$");
            //boolean variable to hold the status
            var isValid = false;
            isValid = !string.IsNullOrEmpty(emlAddress) && myRegex.IsMatch(emlAddress);
            //return the results
            return isValid;
        }

        public ActionResult Sent()
        {
            return View();
        }

        private static string GemplasticsMail(string contactPerson, string tel, string email, string query, string width, string length, string micron, string plastic, string color1, string color2, string color3)
        {
            var sb = new StringBuilder();
            sb.Append(@"<html xmlns='http://www.w3.org/1999/xhtml' >
<head runat='server'>
    <title></title>
    <style type='text/css'>
        .style1
        {
            text-align: left;
        }
    </style>
</head>
<body>
   <table width='100%' border='0' cellspacing='0' cellpadding='0'>
                    <tr>
                        <td align='center' valign='top'>
                            <p>
                                <font color='#000066'><strong>Quotes and Queries</strong></font></p>
                            <table width='85%' border='2' cellspacing='0' cellpadding='1' style='text-align:left'>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td width='54%' class='style1'>
                                        <font color='#FFFFFF'>Contact Person</font>
                                    </td>
                                    <td width='46%' style='text-align: left'>
                                        <input ID='txtContactPerson' value='" + contactPerson + @"' Style='background-color: #6666CC;
                                            color: #FFFFFF' size='35' />
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'>Email Address</font>
                                    </td>
                                    <td style='text-align: left'>
                                        <input ID='txtPreferredContact' value='" + email + @"' Style='background-color: #6666CC;
                                            color: #FFFFFF' size='35' />
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'>Telephone Number</font>
                                    </td>
                                    <td style='text-align: left'>
                                        <input ID='txtTelNum' value='" + tel + @"' Style='background-color: #6666CC; color: #FFFFFF'
                                            size='35' />
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'>Query</font>
                                    </td>
                                    <td style='text-align: left'>
                                        <textarea ID='txtQuery' Style='background-color: #6666CC; color: #FFFFFF; width: 219px;'
                                            size='35' Rows='4' >" + query + @"</textarea>
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'>&nbsp;</font>
                                    </td>
                                    <td class='style1'>
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'><strong>Request Details (non mandatory)</strong></font>
                                    </td>
                                    <td class='style1'>
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'>&nbsp;</font>
                                    </td>
                                    <td class='style1'>
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'>Width</font>
                                    </td>
                                    <td class='style1'>
                                        <input ID='txtWidth' value='" + width + @"' Style='background-color: #6666CC; color: #FFFFFF'
                                            size='35' />
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'>Length</font>
                                    </td>
                                    <td class='style1'>
                                        <input ID='txtLength' value='" + length + @"' Style='background-color: #6666CC; color: #FFFFFF'
                                            size='35' />
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'>Micron(thickness)</font>
                                    </td>
                                    <td class='style1'>
                                        <input ID='txtMicron' value='" + micron + @"' Style='background-color: #6666CC; color: #FFFFFF'
                                            size='35' />
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'>&nbsp;</font>
                                    </td>
                                    <td class='style1'>
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'><strong>Colour of plastic:</strong></font>
                                    </td>
                                    <td class='style1'>
                                        <input ID='txtPlasticColor' value='" + plastic + @"' Style='background-color: #6666CC;
                                            color: #FFFFFF' size='35' />
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'><strong>Print Colors:</strong></font>
                                    </td>
                                    <td class='style1'>
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'>Color One</font>
                                    </td>
                                    <td class='style1'>
                                        <input ID='txtColorOne' value='" + color1 + @"' Style='background-color: #6666CC; color: #FFFFFF'
                                            size='35' />
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'>Color Two</font>
                                    </td>
                                    <td class='style1'>
                                        <input ID='txtColorTwo' value='" + color2 + @"' Style='background-color: #6666CC; color: #FFFFFF'
                                            size='35' />
                                    </td>
                                </tr>
                                <tr bordercolor='#999999' bgcolor='#333399'>
                                    <td class='style1'>
                                        <font color='#FFFFFF'>Color Three</font>
                                    </td>
                                    <td class='style1'>
                                        <input ID='txtColorThree' value='" + color3 + @"' Style='background-color: #6666CC;
                                            color: #FFFFFF' size='35' />
                                    </td>
                                </tr>
                                </table>
                        </td>
                    </tr>
                </table>
</body>
</html>");
            return sb.ToString().Replace("'", "\"");
        }

        private static MailMessage CreateMailMessage(string message, string recipient, string from, string subject)
        {
            var mailMessage = new MailMessage(from, recipient)
            {
                Body = message,
                Subject = subject,
                Priority = MailPriority.High,
                IsBodyHtml = true
            };
            return mailMessage;
        }

        private static SmtpClient CreateSmtpClient(string server, int port, string username, string password)
        {
            var smtpClient = new SmtpClient(server, port);
            var credentials = new NetworkCredential(username, password);
            smtpClient.Credentials = credentials;
            smtpClient.EnableSsl = true;
            smtpClient.Timeout = 20000;
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;

            return smtpClient;
        }

        private static void SendRequest(string message, string recipientEmail, string from, string subject)
        {
            SmtpClient smtpClient = CreateSmtpClient
               (
                    ConfigurationManager.AppSettings["smtpServer"],
                    Convert.ToInt32(ConfigurationManager.AppSettings["port"]),
                    ConfigurationManager.AppSettings["username"],
                    ConfigurationManager.AppSettings["password"]
                );

            MailMessage mailMessage = CreateMailMessage(message, recipientEmail, from, subject);

            try
            {
                smtpClient.Send(mailMessage);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

    }
}
