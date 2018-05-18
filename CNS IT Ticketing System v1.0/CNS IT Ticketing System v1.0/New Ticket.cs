using System;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.DirectoryServices.AccountManagement;
using System.Threading;
using System.ComponentModel;

namespace CNS_IT_Ticketing_System_v1._0
{
    public partial class Form2 : Form
    {
        Form1 _owner;
        bool isadmin;
        public Form2(Form1 owner, bool isadmini)
        {
            InitializeComponent();
            _owner = owner;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form2_FormClosing);
            isadmin = isadmini;

            //// To report progress from the background worker we need to set this property
            //backgroundWorker1.WorkerReportsProgress = true;
            //// This event will be raised on the worker thread when the worker starts
            //backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            //// This event will be raised when we call ReportProgress
            //backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
        }
        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            _owner.LoadData();
        }


        //void Progression(object sender, EventArgs e)
        //{
        //    // Start the background worker
        //    backgroundWorker1.RunWorkerAsync();
        //}
        //// On worker thread so do our thing!
        //void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        //{

        //}
        //// Back on the 'UI' thread so we can update the progress bar
        //void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        //{
        //    // The progress percentage is a property of e
        //    progressBar1.Value = e.ProgressPercentage*2;
        //}



        //loading of data
        private void Form2_Load(object sender, EventArgs e)
        {
            //panel1.Visible = false;
            label10.Text = "";
            label11.Visible = false;
            label12.Visible = false;
            label13.Visible = false;
            this.ActiveControl = textBox2;
            this.textBox1.Hide();
            this.textBox4.Hide();
            this.textBox5.Hide();
            this.textBox6.Hide();
            this.textBox7.Hide();
            this.label15.Visible = false;

            Someoneelse.Checked = false;
            Someoneelse.Visible = false;
            Someoneelselabel.Visible = false;
            comboBox2.Visible = false;
            comboBox2.Enabled = false;

            //Assign ticket to someone
            if (isadmin)
            {
                Someoneelse.Visible = true;

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
                            comboBox2.Items.Add(p.DisplayName);
                    }
                    grp.Dispose();
                    ctx.Dispose();
                }
                catch
                {
                    MessageBox.Show("We are sorry, we are not able to run the program at this time. Please check Internet and VPN connections.");
                    _owner.Close();
                }
                comboBox2.Sorted = true;
            }

            username.Text = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            try
            {
                string domainName = "192.168.10.5";

                //get AD users
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domainName);
                UserPrincipal user = UserPrincipal.FindByIdentity(ctx, username.Text.ToString());
                textBox2.Text = user.EmailAddress;
                textBox3.Text = user.VoiceTelephoneNumber;
            }
            catch
            { }


            DateTime mountain = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.FindSystemTimeZoneById("Mountain Standard Time"));
            label5.Text = mountain.ToString("MM-dd-yyyy h:mm tt");
            label2.Text = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            label4.Text = mountain.ToString("MM-dd-yyyy h:mm tt");
            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
            timer.Interval = (900 * 1000);
            timer.Tick += new EventHandler(timer_tick);
            timer.Start();


        }
        //closes form if no information entered after 15 minutes. 
        private void timer_tick(object sender, EventArgs e)
        {
            this.Close();
        }

        //form close
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //returns a string of char ' ' that is length sized. 
        public string chars(int length)
        {
            string temp = "";
            for (int i = 0; i < 100 - length; i++)
            {
                temp += "-";
            }
            return temp;
        }

        //Available ticket types
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
                if (comboBox1.SelectedIndex == 0)
                {
                    this.textBox1.Show();
                    this.textBox4.Show();
                    this.textBox5.Show();
                    this.textBox6.Show();
                    this.textBox7.Show();
                    this.label11.Visible = true;
                    this.label12.Visible = true;
                    this.label13.Visible = true;
                    this.label15.Visible = true;
                    this.label17.Visible = false;


                    label8.Text = "AD";
                    label10.Text = "Name of Agent needing assistance:";
                    label11.Text = "Username of Agent:";
                    label12.Text = "Best Way to Contact Agent:";
                    label13.Text = "Agent Contact Information:";
                }
                if (comboBox1.SelectedIndex == 1)
                {
                    this.textBox1.Show();
                    this.textBox4.Show();
                    this.textBox5.Show();
                    this.textBox6.Hide();
                    this.textBox7.Hide();
                    this.label11.Visible = true;
                    this.label12.Visible = false;
                    this.label13.Visible = false;
                    this.label15.Visible = true;
                    this.label17.Visible = false;

                    label8.Text = "Desktop";
                    label10.Text = "Error (If Any)";
                    label11.Text = "CNS-ID:";
                }
                if (comboBox1.SelectedIndex == 2)
                {
                    this.textBox1.Show();
                    this.textBox4.Show();
                    this.textBox5.Hide();
                    this.textBox6.Hide();
                    this.textBox7.Hide();
                    this.label11.Visible = false;
                    this.label12.Visible = false;
                    this.label13.Visible = false;
                    this.label15.Visible = true;
                    this.label17.Visible = false;

                    label8.Text = "Email";
                    label10.Text = "Error (If Any)";
                }
                if (comboBox1.SelectedIndex == 3)
                {
                    this.textBox1.Show();
                    this.textBox4.Show();
                    this.textBox5.Show();
                    this.textBox6.Hide();
                    this.textBox7.Hide();
                    this.label11.Visible = true;
                    this.label12.Visible = false;
                    this.label13.Visible = false;
                    this.label15.Visible = true;
                    this.label17.Visible = false;

                    label8.Text = "Fax";
                    label10.Text = "Incoming or outgoing Fax:";
                    label11.Text = "What was the error:";
                }
                if (comboBox1.SelectedIndex == 4)
                {
                    this.textBox1.Show();
                    this.textBox4.Show();
                    this.textBox5.Show();
                    this.textBox6.Show();
                    this.textBox7.Show();
                    this.label11.Visible = true;
                    this.label12.Visible = true;
                    this.label13.Visible = true;
                    this.label15.Visible = true;
                    this.label17.Visible = false;


                    label10.Text = "Name of Agent needing reset:";
                    label11.Text = "Username of Agent:";
                    label12.Text = "Best Way to Contact Agent:";
                    label13.Text = "Agent Contact Information:";
                    label8.Text = "Kinnser Password";
                }
                if (comboBox1.SelectedIndex == 5)
                {
                    this.textBox1.Show();
                    this.textBox4.Show();
                    this.textBox5.Show();
                    this.textBox6.Hide();
                    this.textBox7.Hide();
                    this.label11.Visible = true;
                    this.label12.Visible = false;
                    this.label13.Visible = false;
                    this.label15.Visible = true;
                    this.label17.Visible = false;

                    label8.Text = "Laptop";
                    label10.Text = "Error (If Any)";
                    label11.Text = "CNS-ID:";
                }
                if (comboBox1.SelectedIndex == 6)
                {
                    this.textBox1.Show();
                    this.textBox4.Show();
                    this.textBox5.Show();
                    this.textBox6.Hide();
                    this.textBox7.Hide();
                    this.label11.Visible = true;
                    this.label12.Visible = false;
                    this.label13.Visible = false;
                    this.label15.Visible = true;
                    this.label17.Visible = false;


                    label8.Text = "Phone";
                    label10.Text = "Type of Phone: ";
                    label11.Text = "Errors (if Any): ";
                }
                if (comboBox1.SelectedIndex == 7)
                {

                    this.textBox1.Show();
                    this.textBox4.Show();
                    this.textBox5.Show();
                    this.textBox6.Hide();
                    this.textBox7.Hide();
                    this.label11.Visible = true;
                    this.label12.Visible = false;
                    this.label13.Visible = false;
                    this.label15.Visible = true;
                    this.label17.Visible = false;


                    label8.Text = "Printer-Scanner";
                    label10.Text = "Who owns the printer:";
                    label11.Text = "Steps you have tried to fix issue:";
                }
                if (comboBox1.SelectedIndex == 8)
                {
                    this.textBox1.Show();
                    this.textBox4.Hide();
                    this.textBox5.Hide();
                    this.textBox6.Hide();
                    this.textBox7.Hide();
                    this.label10.Text = "";
                    this.label11.Visible = false;
                    this.label12.Visible = false;
                    this.label13.Visible = false;
                    this.label15.Visible = true;
                    this.label17.Visible = false;

                    label8.Text = "Other";
                }
                if (comboBox1.SelectedIndex == 9)
                {
                    this.textBox1.Show();
                    this.textBox4.Show();
                    this.textBox5.Show();
                    this.textBox6.Hide();
                    this.textBox7.Hide();
                    
                    this.label11.Visible = true;
                    this.label12.Visible = false;
                    this.label13.Visible = false;
                    this.label15.Visible = true;
                    this.label17.Visible = false;

                    label8.Text = "Tablet";
                    label10.Text = "Error (If Any)";
                    label11.Text = "Contact Phone Number:";
                }
                if (comboBox1.SelectedIndex == 10)
                {
                    this.textBox1.Show();
                    this.textBox4.Show();
                    this.textBox5.Hide();
                    this.textBox6.Hide();
                    this.textBox7.Hide();
                    this.label11.Visible = false;
                    this.label12.Visible = false;
                    this.label13.Visible = false;
                    this.label15.Visible = true;
                    this.label17.Visible = false;

                    label8.Text = "CRM Dynamics 365";
                    label10.Text = "Error (If Any)";
                }
                if(comboBox1.SelectedIndex == 11)
                {
                    this.textBox1.Hide();
                    this.textBox4.Hide();
                    this.textBox5.Hide();
                    this.textBox6.Hide();
                    this.textBox7.Hide();
                    this.label11.Visible = false;
                    this.label12.Visible = false;
                    this.label13.Visible = false;
                    this.label15.Visible = false;
                    this.label17.Visible = false;
                    this.label8.Text = "Kinnser";
                    label10.Text = "Please clear your History (Cookies and Cache). \r\nIf this does not work, " +
                    "please contact Kinnser support at support@kinnser.com or by calling (877) 399-6538. Thank you.";
                }
        }

        //submit new ticket
        private void button1_Click(object sender, EventArgs e)
        {            
            if (comboBox1.SelectedIndex == -1)
            {
                
                MessageBox.Show("User MUST select a ticket type", "Error, invalid selection",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (!textBox2.Text.ToString().Contains("@") || !textBox2.Text.ToString().Contains(".com"))
                    MessageBox.Show("Please Enter a Valid email address");
                else if (!textBox2.Text.ToString().Contains("cnscares.com"))
                {
                    MessageBox.Show("Please use your CNS email. No other email address can be accepted.");
                }
                else
                {
                    Random rnd = new Random();

                    DateTime mountain = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.FindSystemTimeZoneById("Mountain Standard Time"));
                    long ticketnumber = long.Parse(mountain.ToString("MMddyyyyhhmm") + rnd.Next(1, 1000).ToString());
                    string ticketor = "";
                    string tphone = "";
                    string temail = "";
                    string dissue = "";
                    string issue = "";
                    char status = 'P';
                    string Device = "";
                    ticketor = username.Text;
                    tphone = textBox3.Text;
                    temail = textBox2.Text;
                    dissue = label5.Text;
                    if (label8.Text == "Fax")
                        issue = "Type of issue: " + label8.Text + "\r\n\r\nIncoming or Outgoing Fax: " + textBox4.Text +
                            "\r\n\r\nError Provided: " + textBox5.Text + "\r\n\r\nMore Details: \r\n" + textBox1.Text;
                    else if (label8.Text == "Kinnser Password" || label8.Text == "AD")
                        issue = "Type of issue: " + label8.Text + "\r\n\r\nName of Agent: " + textBox4.Text +
                            "\r\n\r\nUsername of Agent: " + textBox5.Text + "\r\n\r\nBest way to contact: " +
                            textBox6.Text + "\r\n\r\nContact Information: " + textBox7.Text + "\r\n\r\nMore Details: \r\n"
                            + textBox1.Text;
                    else if (label8.Text == "Printer-Scanner")
                        issue = "Type of issue: " + label8.Text + "\r\n\r\nOwner of Printer-Scanner: " + textBox4.Text +
                            "\r\n\r\nPreviously completed Steps: " + textBox5.Text + "\r\n\r\nMore Details: \r\n" + textBox1.Text;
                    else if (label8.Text == "Email" || label8.Text == "Laptop" || label8.Text == "Desktop" ||
                        label8.Text == "Tablet")
                    {
                        if (label8.Text == "Laptop" || label8.Text == "Desktop")
                        {
                            Device = textBox5.Text;
                        }
                        issue = "Type of issue: " + label8.Text + "\r\n\r\nError: " + textBox4.Text +
                            "\r\n\r\nMore Details: \r\n" + textBox1.Text;
                    }
                    else if (label8.Text == "Phone")
                        issue = "Type of issue: " + label8.Text + "\r\n\r\nType of Phone: " + textBox4.Text +
                            "\r\n\r\nErrors: " + textBox5.Text + "\r\n\r\nMore Details: \r\n" + textBox1.Text;
                    else if (label8.Text == "CRM Dynamics 365")
                    {
                        issue = "Type of issue: D365\r\n\r\nError: " + textBox4.Text + "\r\n\r\nMore Details: \r\n" + textBox1.Text;
                        label8.Text = "D365";
                    }
                    else
                        issue = textBox1.Text;
                    if (label8.Text != "Kinnser")
                    {
                        if (label8.Text == "Kinnser Password")
                            label8.Text = "Kinnser";
                        string notes = "";
                        if (Someoneelse.Checked)
                        {
                            string domainName = "192.168.10.5";
                            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domainName);

                            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, comboBox2.Text);

                            if (user != null)
                            {
                                notes += "Ticket was created by: " + ticketor + "\r\nFor: CNS\\" + user.SamAccountName + "\r\n\r\n";
                                ticketor = "CNS\\" + user.SamAccountName;
                            }
                            else
                            {

                                MessageBox.Show("Error! No selection made that is in the server!");
                            }
                        }
                        using (TicketsDBDataContext dbContext = new TicketsDBDataContext())
                        {
                            Ticketer test = new Ticketer
                            {
                                ID2 = ticketnumber,
                                Ticketor = ticketor,
                                Tphone = tphone,
                                Temail = temail,
                                Dissue = dissue,
                                Issue = issue,
                                Status = status,
                                DeviceID = Device,
                                IssueType = label8.Text,
                                Resolver = null,
                                Priority = 'P',
                                Supported = true,
                                Assigned = "It Support",
                                Notes = notes
                            };
                            dbContext.Ticketers.InsertOnSubmit(test);
                            bool submitsuccess = false;
                            if (label8.Text == "Tablet" && textBox5.Text == "")
                            {
                                MessageBox.Show("A contact phone number is required");
                            }
                            else
                            {
                                try
                                {
                                    dbContext.SubmitChanges();
                                    submitsuccess = true;
                                }
                                catch
                                {
                                    submitsuccess = false;
                                    MessageBox.Show("Unable to create a new ticket. No server connection available. Please check VPN and internet connections");
                                    _owner.Close();
                                }


                            if(submitsuccess==true)
                            {
                                //emailcode here
                                try
                                {
                                    //standard email for a D365 ticket
                                    if (label8.Text == "D365")
                                    {
                                        email D365email = new email();
                                        D365email.sendemail("A new D365 IT Ticket has been placed.", "A new IT ticket for D365 related issues has been created. please review the details of this ticket in the CNS IT ticketing system.\r\n\r\nTicket ID: "+ticketnumber+"\r\nType of issue: D365", ticketnumber.ToString(), "D365support@cnscares.com");
                                    }

                                    //Standard email for a new ticket.
                                    email NewTicket = new email();
                                    NewTicket.sendemail("Your IT ticket has been created.", "Thank you for creating your IT ticket using the CNS IT ticketing system.\r\n\r\nThe ticket ID is: " + ticketnumber + "\r\nThe type of issue is: " + label8.Text + "\r\n\r\nThis issue will be resolved as quickly as possible. Please allow 10 minutes for the IT team to see this new ticket.\r\n\r\nThank you,\r\nYour IT Support Team\r\n\r\n", ticketnumber.ToString(), textBox2.Text.ToString());

                                    MessageBox.Show("Your ticket is submited. The ID is: " + ticketnumber, "Ticket Creation Successful!",
MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                catch(Exception ex)
                                {
                                    MessageBox.Show("Unable to send email(s).\r\n\r\n Error: \r\n" + ex.ToString());
                                }
                                this.Close();
                            }
                            else
                            {
                                this.Close();
                            }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Sorry, this is not a supported ticket type.");
                        this.Close();
                    }
                }
            }
        }

        private void Someoneelse_CheckedChanged(object sender, EventArgs e)
        {
            if(Someoneelse.Checked)
            {
                comboBox2.Enabled = true;
                comboBox2.Visible = true;
                Someoneelselabel.Visible = true;
                label14.Text = "Their Email:";
                label16.Text = "Their Phone Number";
            }
            else
            {
                comboBox2.Enabled = false;
                comboBox2.Visible = false;
                Someoneelselabel.Visible = false;
                label14.Text = "Your Email:";
                label16.Text = "Your Phone Number";
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string username = comboBox2.Text;
            string domainName = "192.168.10.5";

            PrincipalContext domainContext = new PrincipalContext(ContextType.Domain, domainName);
            UserPrincipal user = UserPrincipal.FindByIdentity(domainContext, username);
            textBox2.Text = user.EmailAddress.ToString();
            textBox3.Text = user.VoiceTelephoneNumber;
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
    }
}
