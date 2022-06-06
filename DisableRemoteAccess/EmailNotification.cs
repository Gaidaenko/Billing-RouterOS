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
                MailAddress to = new MailAddress("recipient@gmail.com");
                MailMessage message = new MailMessage(from, to);

                message.IsBodyHtml = true;
                message.Subject = "Доступ к серверу " + Fields.customerName + "включен!";
                message.Body = "Здравствуйте. Сервер " + Fields.customerName + " возобновил работу.";
                               
                SmtpClient smtp = new SmtpClient("smtp.mail.yahoo.com", 587);
                smtp.Credentials = new NetworkCredential("mail@yahoo.com", "password");
                smtp.EnableSsl = true;
                smtp.Send(message);
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
                MailAddress to = new MailAddress("recipient@gmail.com");
                MailMessage message = new MailMessage(from, to);

                message.IsBodyHtml = true;
                message.Subject = "Доступ к серверу "+ Fields.customerName + " отключен!";
                message.Body = "Здравствуйте. Сервер " + Fields.customerName + " был отключен отключен за неуплату. " +
                              "<br/>" +
                               "<br/>";

                SmtpClient smtp = new SmtpClient("smtp.mail.yahoo.com", 587);
                smtp.Credentials = new NetworkCredential("mail@yahoo.com", "password");
                smtp.EnableSsl = true;
                smtp.Send(message);
                return;
            }
            catch (Exception e)
            {



                return;
            }
        }
    }
}
