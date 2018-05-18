using System;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.DirectoryServices.AccountManagement;


namespace CNS_IT_Ticketing_System_v1._0
{
    public partial class EscalateReturn : Form
    {
        string RID = "";
        string eemail = "";
        ReturnData owner;
        public EscalateReturn(string ID, ReturnData owner_)
        {
            InitializeComponent();
            RID = ID;
            owner = owner_;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (EquipReDataContext dbContext = new EquipReDataContext())
            {
                try
                {
                    EquipmentReturn submitchanges = dbContext.EquipmentReturns.SingleOrDefault(X => X.IDText.Contains(RID));
                    submitchanges.EscalationAgent = comboBox1.Text;
                    submitchanges.EscalationStatus = true;
                    submitchanges.EscalationDate = DateTime.Now;
                    submitchanges.Notes += "\r\n\r\nEmail of Escalation Supervisor: " + textBox2.Text;

                    eemail = textBox2.Text;

                    submitchanges.TicketStatus = 'C';

                    sendemail(submitchanges.TagID, textBox2.Text, submitchanges.Phone_Number.ToString(), submitchanges.Assignee.ToString(), comboBox1.Text);

                    dbContext.SubmitChanges();
                    owner.Close();
                    this.Close();
                }
                catch(Exception explanation)
                {
                    MessageBox.Show("Error! Unable to connect to server\r\n\r\n"+ explanation);
                }
            }

        }

        private void sendemail(string TagID, string eto, String phonenumber, string NameOfAssignee, string ManagerName)
        {
            try
            {
                using (SmtpClient smtp = new SmtpClient())
                {
                    smtp.Host = "mail.cnscares.com";
                    smtp.UseDefaultCredentials = false;
                    NetworkCredential netCred = new NetworkCredential("CNSIT.ITTickets@cnscares.com", "Christmas1a", "CNS");
                    smtp.Credentials = netCred;
                    smtp.EnableSsl = false;
                    using (MailMessage msg = new MailMessage("CNSIT.Tickets@cnscares.com", eto.ToString()))
                    {
                        msg.Subject = "Escalated Ticket";
                        StringBuilder sb = new StringBuilder();
                        sb.AppendLine("This is an escalated email message regarding the return of equipment to the corporate office, three previous attempts were made to contact this individual. Please contact this individual and instruct them to return the equipment to their local CNS office or arrange a return shipping box that will be sent to their home address. Return boxes can be requested by calling 877-259-9001. Please send an email to itsupport@cnscares.com within 48 hours with the updated recovery plan. \r\n\r\nDevice ID: " + TagID+"\r\nPhone Number Called: "+phonenumber+"\r\nDevice Assigned to: "+NameOfAssignee+ "\r\n\r\n\r\n");
                        sb.AppendLine("\r\n\r\nPlease note: This is an automated Message. Any replies to this email will be forwarded to itsupport@cnscares.com");
                        msg.Body = sb.ToString();
                        msg.IsBodyHtml = false;
                        smtp.Send(msg);
                    }
                }
                using (SmtpClient smtp = new SmtpClient())
                {
                    smtp.Host = "mail.cnscares.com";
                    smtp.UseDefaultCredentials = false;
                    NetworkCredential netCred = new NetworkCredential("CNSIT.ITTickets@cnscares.com", "Christmas1a", "CNS");
                    smtp.Credentials = netCred;
                    smtp.EnableSsl = false;
                    using (MailMessage msg = new MailMessage("CNSIT.Tickets@cnscares.com", "EquipmentReturns@cnscares.com"))
                    {
                        msg.Subject = "Escalated Ticket";
                        StringBuilder sb = new StringBuilder();
                        sb.AppendLine("This is an escalated email message regarding the return of equipment to the corporate office, three previous attempts were made to contact this individual. The Manager of the Assignee was emailed regarding this. \r\n\r\nDevice ID: " + TagID + "\r\nPhone Number Called: " + phonenumber + "\r\nDevice Assigned to: " + NameOfAssignee + "\r\nManager's Name: "+ManagerName+"\r\n\r\n\r\n");
                        sb.AppendLine("\r\n\r\nPlease note: This is an automated Message. Any replies to this email will be forwarded to itsupport@cnscares.com");
                        msg.Body = sb.ToString();
                        msg.IsBodyHtml = false;
                        smtp.Send(msg);
                    }
                }
                try
                {
                    MessageBox.Show("Ticket Has been Escalated", "Ticket Escalation Successful!",
MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch
                {
                    MessageBox.Show("Unable to create a new ticket. No server connection available");
                }
                this.Close();
            }
            catch
            {
                MessageBox.Show("Please enter a valid email address.");
            }
        }

        private void EscalateReturn_Load(object sender, EventArgs e)
        {
            string groupName = "Domain Users";
            string domainName = "192.168.10.5";

            //get AD users
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domainName);
            GroupPrincipal grp = GroupPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, groupName);

            try
            {
                foreach (Principal p in grp.GetMembers(false))
                {
                    if (p.DisplayName != null)
                        comboBox1.Items.Add(p.DisplayName);
                }
                grp.Dispose();
                ctx.Dispose();
            }
            catch
            {
                MessageBox.Show("Server not available. Check internet connection");
            }
            comboBox1.Sorted = true;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string username = comboBox1.Text;
            string domainName = "192.168.10.5";

            PrincipalContext domainContext = new PrincipalContext(ContextType.Domain, domainName);
            UserPrincipal user = UserPrincipal.FindByIdentity(domainContext, username);
            textBox2.Text = user.EmailAddress.ToString();
        }
    }
}
