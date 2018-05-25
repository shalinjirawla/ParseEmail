using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using EAGetMail;


namespace DailySendEmail
{
   public class SendMailService
    {
        // This function write log to LogFile.text when some error occurs.      
        public static void WriteErrorLog(Exception ex)
        {
            StreamWriter sw = null;
            try
            {
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + ex.Source.ToString().Trim() + "; " + ex.Message.ToString().Trim());
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }
        // This function write Message to log file.    
        public static void WriteErrorLog(string Message)
        {
            StreamWriter sw = null;
            try
            {
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + Message);
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }
        // This function contains the logic to send mail.    
        public static void SendEmail()
        {
                MailServer oServer = new MailServer("imap.gmail.com",
                       "haresh.ncoresoft@gmail.com", "Ncoresoft@123", ServerProtocol.Imap4);
                MailClient oClient = new MailClient("TryIt");

                oServer.SSLConnection = true;

                oServer.Port = 993;

            try
            {
                oClient.Connect(oServer);
                MailInfo[] infos = oClient.GetMailInfos();
                for (int i = infos.Length; i > 0; i--)
                {
                    if (infos[i - 1].Read == false)
                    {
                        MailInfo info = infos[i - 1];
                       
                        Mail oMail = oClient.GetMail(info);

                        var body = "";

                        body = (oMail.TextBody);
                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml(body);
                        var fullname = "";
                        foreach (XmlNode node in doc.GetElementsByTagName("name"))
                        {
                            var part = node.Attributes["part"];
                            if (part != null)
                            {
                                if (part.Value == "full")
                                {
                                    fullname = node.InnerText;
                                }
                            }
                        }

                        foreach (XmlNode node in doc.GetElementsByTagName("vendorname"))
                        {
                            node.InnerText = node.InnerText + "[Attn: " + fullname + "]";
                        }

                        System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                        mail.To.Add(ConfigurationManager.AppSettings["EmailTo"]);
                        mail.From = new System.Net.Mail.MailAddress("rajpadaliya1996@gmail.com");
                        mail.Subject = "new Xml";
                        mail.Body = doc.InnerXml;
                        SmtpClient smtp = new SmtpClient
                        {
                            Host = "smtp.gmail.com",
                            Port = 587,
                            UseDefaultCredentials = false,
                            Credentials = new System.Net.NetworkCredential("nikunjparmar25594@gmail.com", "@honor6x"),
                            EnableSsl = true
                        };
                        smtp.Send(mail);
                    

                    oClient.MarkAsRead(info, true);

                    }
                    else
                    {
                        oClient.Quit();
                    }
                }
             }
            catch (Exception ex)
            {
                WriteErrorLog(ex.InnerException.Message);
                throw;
            }
        }
    }
}
