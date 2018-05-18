using System;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.DirectoryServices.AccountManagement;

namespace CNS_IT_Ticketing_System_v1._0
{
    public partial class CloseTicket : Form
    {
        Form1 _owner;
        long useselected;
        public CloseTicket(long selectedID, Form1 owner)
        {
            InitializeComponent();
            useselected = selectedID;
            _owner = owner;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.CloseTicket_FormClosing);
        }
        private void CloseTicket_FormClosing(object sender, FormClosingEventArgs e)
        {
            _owner.LoadData();
        }
        private void CloseTicket_Load(object sender, EventArgs e)
        {
            username.Text = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            label10.Text = DateTime.Now.ToString("MM-dd-yyyy h:mm tt");
            string ID = "";
            ID += useselected;
            label3.Text = ID;

            try
            {
                string domainName = "192.168.10.5";

                //get AD users
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domainName);
                UserPrincipal user = UserPrincipal.FindByIdentity(ctx, username.Text.ToString());
                textBox4.Text = user.EmailAddress;
                textBox5.Text = user.VoiceTelephoneNumber;
            }
            catch
            { }
        }

        //submit close of ticket.
        private void button1_Click(object sender, EventArgs e)
        {
            using (TicketsDBDataContext dbContext = new TicketsDBDataContext())
            {
                Ticketer closer = dbContext.Ticketers.SingleOrDefault(X => X.ID2 == useselected);
                closer.Notes += textBox2.Text;
                if (textBox1.Text == "" || textBox4.Text == "" || textBox5.Text == "")
                {
                    MessageBox.Show("Please fill in all details to close the ticket", "Status update unsuccessful",
MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    closer.Dresolve = label10.Text;
                    closer.Resolver = username.Text;
                    closer.Remail = textBox4.Text;
                    closer.Rphone = textBox5.Text;
                    closer.Resolution = textBox1.Text;
                    closer.Status = 'C';
                    bool isadmin = false;
                    if (closer.IssueType == "IT-Update"|| closer.IssueType == "IT-New PC"
                        || closer.IssueType == "IT-Permissions" || closer.IssueType == "IT-Destroy"
                        || closer.IssueType == "IT-Phone" || closer.IssueType == "IT-Security"
                        || closer.IssueType == "IT-Tablet" || closer.IssueType == "IT-User"
                        || closer.IssueType == "IT-Update"|| closer.IssueType == "IT-Other"
                        || closer.IssueType=="IT-ODI")
                        isadmin = true;

                    sendemail(useselected, closer.Temail, dbContext, isadmin);
                    MessageBox.Show("Update Successful");
                    this.Close();
                }
            }
        }
        private void sendemail(long ticket, string eto, TicketsDBDataContext dbContext, bool isadmin)
        {
            if (isadmin == false)
            {
                try
                {
                    email ClosedTicket = new email();
                    ClosedTicket.sendemail("Your IT Ticket has been closed.", "Thank you for patiently waiting while we worked to resolve your issue.\r\n\r\nYour resolved ticket ID is: " + ticket+ "\r\n\r\nThank you,\r\nYour IT Support Team", ticket.ToString(), eto.ToString());
                    try
                    {
                        dbContext.SubmitChanges();
                    }
                    catch
                    {
                        MessageBox.Show("Unable to create a new ticket. No server connection available");
                    }
                    this.Close();
                }
                catch
                {
                    MessageBox.Show("User email address not valid. Saved changes and closed ticket without sending email.");
                    dbContext.SubmitChanges();
                    this.Close();
                }
            }
            else
            {
                try
                {
                    dbContext.SubmitChanges();
                }
                catch
                {
                    MessageBox.Show("Unable to create a new ticket. No server connection available");
                }
                 this.Close();
            }
        }

        //cancel close of ticket.
        private void button2_Click(object sender, EventArgs e)
        {

            this.Close();
        }
    }
}
