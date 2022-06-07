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
                MailAddress toTheClient = new MailAddress(Fields.addrMail);
                MailMessage messageToTheClient = new MailMessage(from, toTheClient);
                MailAddress bcc = new MailAddress("recipient@gmail.com");
                messageToTheClient.Bcc.Add(bcc);

                messageToTheClient.IsBodyHtml = true;
                messageToTheClient.Subject = "Доступ к серверу " + Fields.customerName + " включен!";
                messageToTheClient.Body = "Здравствуйте. Сервер " + Fields.customerName + " возобновил работу.";
                
                SmtpClient smtp = new SmtpClient("smtp.mail.yahoo.com", 587);
                smtp.Credentials = new NetworkCredential("mail@yahoo.com", "password");
                smtp.EnableSsl = true;
                smtp.Send(messageToTheClient);
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
                MailAddress toTheClient = new MailAddress(Fields.addrMail);
                MailMessage messageToTheClient = new MailMessage(from, toTheClient);
                MailAddress bcc = new MailAddress("recipient@gmail.com");
                messageToTheClient.Bcc.Add(bcc);

                messageToTheClient.IsBodyHtml = true;
                messageToTheClient.Subject = "Доступ к серверу "+ Fields.customerName + " отключен!";
                messageToTheClient.Body = "Здравствуйте. Сервер " + Fields.customerName + " был отключен за неуплату. " +
                              "<br/>"+
                              "<br/>";

                SmtpClient smtp = new SmtpClient("smtp.mail.yahoo.com", 587);
                smtp.Credentials = new NetworkCredential("mail@yahoo.com", "password");
                smtp.EnableSsl = true;
                smtp.Send(messageToTheClient);
                return;
            }
            catch (Exception e)
            {



                return;
            }
        }
    }
}
