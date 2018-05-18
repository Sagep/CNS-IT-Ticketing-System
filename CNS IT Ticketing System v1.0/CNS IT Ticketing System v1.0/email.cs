using System.Net.Mail;
using System.Net;

namespace CNS_IT_Ticketing_System_v1._0
{
    public class email
    {
        private SmtpClient smtp = new SmtpClient();
        private NetworkCredential netCred = new NetworkCredential("CNSIT.ITTickets@cnscares.com", "Christmas1a", "CNS");

        private string subject;
        private string body;
        private string ticketID;
        private string targetemail;
        private string disclaimer = "Please note: This is an automatically generated email. Any replies to this email will be forwarded to itsupport@cnscares.com";
        private MailMessage msg;
        public email()
        {
            //setup of email information
            smtp.Host = "mail.cnscares.com";
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = netCred;
            smtp.EnableSsl = false;
        }

        public void sendemail(string SB, string BD, string TID, string target)
        {
            targetemail = target;
            subject = SB;
            body = BD;
            ticketID = TID;
            send();
        }
        private void send()
        {
            msg = new MailMessage("CNSIT.Tickets@cnscares.com", targetemail);
            msg.Subject = subject + " Ticket ID: " + ticketID;
            msg.Body = body + "\r\n\r\n" + disclaimer;
            msg.IsBodyHtml = false;
            smtp.Send(msg);
        }
    }

}
