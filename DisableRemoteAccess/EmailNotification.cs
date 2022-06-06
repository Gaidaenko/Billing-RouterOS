using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace DisableRemoteAccess
{
    public static class EmailNotification
    {
        public static void messageAccessEnabled()
        {
            try
            {
                MailAddress from = new MailAddress("mail@yahoo.com", "Company");
                MailAddress toTheHead = new MailAddress("recipient@gmail.com");
                MailMessage messageToTheHead = new MailMessage(from, toTheHead);
                MailAddress bcc = new MailAddress("recipient2@gmail.com");
                messageToTheHead.CC.Add(bcc);

                messageToTheHead.IsBodyHtml = true;
                messageToTheHead.Subject = "Доступ к серверу " + Fields.customerName + " включен!";
                messageToTheHead.Body = "Здравствуйте. Сервер " + Fields.customerName + " возобновил работу.";
                               
                SmtpClient smtp = new SmtpClient("smtp.mail.yahoo.com", 587);
                smtp.Credentials = new NetworkCredential("mail@yahoo.com", "password");
                smtp.EnableSsl = true;
                smtp.Send(messageToTheHead);
               // smtp.Send(messageToTheHead2);
                return;
            }
            catch (Exception e)
            {


                return;
            }
        }
        public static void messageAccessDisabled()
        { 
            try
            {
                MailAddress from = new MailAddress("mail@yahoo.com", "Company");
                MailAddress toTheHead = new MailAddress("recipient@gmail.com");
                MailMessage messageToTheHead = new MailMessage(from, toTheHead);
                MailAddress bcc = new MailAddress("recipient2@gmail.com");
                messageToTheHead.CC.Add(bcc);

                messageToTheHead.IsBodyHtml = true;
                messageToTheHead.Subject = "Доступ к серверу "+ Fields.customerName + " отключен!";
                messageToTheHead.Body = "Здравствуйте. Сервер " + Fields.customerName + " был отключен отключен за неуплату. " +
                              "<br/>" +
                               "<br/>";

                SmtpClient smtp = new SmtpClient("smtp.mail.yahoo.com", 587);
                smtp.Credentials = new NetworkCredential("mail@yahoo.com", "password");
                smtp.EnableSsl = true;
                smtp.Send(messageToTheHead);
                return;
            }
            catch (Exception e)
            {



                return;
            }
        }
    }
}
