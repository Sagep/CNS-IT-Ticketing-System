using System;
using System.Data;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Windows.Forms;


namespace CNS_IT_Ticketing_System_v1._0
{
    public partial class UpdateTicket : Form
    {
        Form1 _owner;
        long useselected;
        bool clicknotes = false;
        bool isadmin;
        public UpdateTicket(long selectedID, Form1 owner, bool admin)
        {
            InitializeComponent();
            useselected = selectedID;
            _owner = owner;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.UpdateTicket_FormClosing);
            isadmin = admin;
        }
        private void UpdateTicket_FormClosing(object sender, FormClosingEventArgs e)
        {
            _owner.LoadData();
        }

        private void UpdateTicket_Load(object sender, EventArgs e)
        {
            if(!isadmin)
            {
                button3.Visible = false;
            }
            username.Text = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            label10.Text = DateTime.Now.ToString("MM-dd-yyyy h:mm tt");
            textBox2.Visible = false;
            TicketsDBDataContext db = new TicketsDBDataContext();
            label21.Text = "" + useselected;
            var r = from p in db.Ticketers
                    where p.ID2==useselected
                    select p;

            foreach (var x in r)
            {
                label2.Text = x.Ticketor;
                if (x.Temail==null)
                {
                    label7.Text = "None";
                }
                else
                    label7.Text = x.Temail;

                if (x.Tphone == null)
                {
                    label8.Text = "None";
                }
                else
                label8.Text = x.Tphone;
                label4.Text = x.Dissue;
                textBox1.Text = x.Issue;
                textBox3.Text = x.Notes;
                if (x.IssueType == "AD")
                    comboBox1.SelectedIndex = 0;
                else if (x.IssueType == "Desktop")
                    comboBox1.SelectedIndex = 1;
                else if (x.IssueType == "Email")
                    comboBox1.SelectedIndex = 2;
                else if (x.IssueType == "Fax")
                    comboBox1.SelectedIndex = 3;
                else if (x.IssueType == "Kinnser")
                    comboBox1.SelectedIndex = 4;
                else if (x.IssueType == "Laptop")
                    comboBox1.SelectedIndex = 5;
                else if (x.IssueType == "Phone")
                    comboBox1.SelectedIndex = 6;
                else if (x.IssueType == "Printer-Scanner")
                    comboBox1.SelectedIndex = 7;
                else if (x.IssueType == "Other")
                    comboBox1.SelectedIndex = 8;
                else if (x.IssueType == "Tablet")
                    comboBox1.SelectedIndex = 9;
                else if (x.IssueType == "D365")
                    comboBox1.SelectedIndex = 10;
                else
                    comboBox1.Visible = false;
                if (x.Status== 'P')
                {
                    checkBox1.Checked = false;
                    checkBox2.Checked = true;
                    checkBox3.Checked = false;
                    checkBox4.Checked = false;
                }
                else if(x.Status=='C')
                {
                    checkBox1.Checked = false;
                    checkBox2.Checked = false;
                    checkBox3.Checked = true;
                    checkBox4.Checked = false;
                    checkBox1.Enabled = false;
                    checkBox2.Enabled = false;
                    checkBox3.Enabled = false;
                    checkBox4.Enabled = false;
                    textBox3.Visible = true;
                    textBox2.ReadOnly = false;
                    textBox4.Visible = false;
                    textBox5.Visible = false;
                    label15.Visible = false;
                    label17.Visible = false;
                }
                else if (x.Status == 'O')
                {
                    checkBox1.Checked = true;
                    checkBox2.Checked = false;
                    checkBox3.Checked = false;
                    checkBox4.Checked = false;
                }
                else if (x.Status == 'F')
                {
                    checkBox1.Checked = false;
                    checkBox2.Checked = false;
                    checkBox3.Checked = false;
                    checkBox4.Checked = true;
                }
                //priority of ticket
                if(x.Priority=='P'||x.Priority=='L')
                {
                    checkBox7.Checked = false;
                    checkBox6.Checked = false;
                    checkBox5.Checked = true;
                }
                else if(x.Priority=='M')
                {
                    checkBox7.Checked = false;
                    checkBox6.Checked = true;
                    checkBox5.Checked = false;
                }
                else if (x.Priority=='H')
                {
                    checkBox7.Checked = true;
                    checkBox6.Checked = false;
                    checkBox5.Checked = false;
                }
                //Supported by IT
                if (x.Supported == true)
                {
                    checkBox8.Checked = true;
                    checkBox9.Checked = false;
                }
                else
                {
                    checkBox8.Checked = false;
                    checkBox9.Checked = true;
                }
            }

        }

        private void Noteclick(object sender, EventArgs e)
        {
            if (clicknotes==false)
            {
                textBox3.Text = "";
                clicknotes = true;
            }
        }

        #region Checkboxes
        //checkbox check functions
        private void checkheckbox(CheckBox A, CheckBox C, CheckBox B)
        {
            if (A.Checked)
            {
                B.Checked = false;
                C.Checked = false;
                A.Checked = true;
            }
            else if (B.Checked == false && C.Checked == false)
            {
                A.Checked = true;
            }
        }
        private void checkheckbox(CheckBox A, CheckBox C, CheckBox B, CheckBox D)
        {
            if (A.Checked)
            {
                B.Checked = false;
                C.Checked = false;
                D.Checked = false;
                A.Checked = true;
            }
            else if (B.Checked == false && C.Checked == false && D.Checked == false)
            {
                A.Checked = true;
            }
        }
        private void checkheckbox(CheckBox A, CheckBox b)
        {
            if(A.Checked)
            {
                b.Checked = false;
                A.Checked = true;
            }
            else if( !b.Checked)
            {
                A.Checked = true;
            }
        }

        //checkboxes for priority
        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            checkheckbox(checkBox5, checkBox6, checkBox7);
        }
        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            checkheckbox(checkBox6, checkBox5, checkBox7);
        }
        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            checkheckbox(checkBox7, checkBox6, checkBox5);
        }

        //checkboxes for Status
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                checkheckbox(checkBox2, checkBox1, checkBox3, checkBox4);
                textBox2.Visible = false;
                label15.Visible = false;
                label17.Visible = false;
                textBox4.Visible = false;
                textBox5.Visible = false;
                label6.Visible = false;
            }
        }
        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                checkheckbox(checkBox1, checkBox2, checkBox3, checkBox4);

                textBox2.Visible = false;
                label15.Visible = false;
                label17.Visible = false;
                textBox4.Visible = false;
                textBox5.Visible = false;
                label6.Visible = false;
            }
        }
        private void checkBox3_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
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

                textBox2.Visible = true;
                label6.Visible = true;
                label15.Visible = true;
                label17.Visible = true;
                textBox4.Visible = true;
                textBox5.Visible = true;
                checkheckbox(checkBox3, checkBox2, checkBox1, checkBox4);

            }
            else
            {
                textBox4.Text = "";
                textBox5.Text = "";
            }
        }
        private void checkBox4_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                checkheckbox(checkBox4, checkBox2, checkBox3, checkBox1);
                textBox2.Visible = false;

                label6.Visible = false;
                label15.Visible = false;
                label17.Visible = false;
                textBox4.Visible = false;
                textBox5.Visible = false;
            }
        }

        //checkboxes for supported
        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            checkheckbox(checkBox8, checkBox9);
        }
        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            checkheckbox(checkBox9, checkBox8);
        }
        #endregion

        //right click copy command
        private void Copy_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(""+useselected);
        }

        //submition of update
        private void Submit(object sender, EventArgs e)
        {
            using (TicketsDBDataContext dbContext = new TicketsDBDataContext())
            {
                try
                {
                    Ticketer test = dbContext.Ticketers.SingleOrDefault(X => X.ID2 == useselected);
                    if (comboBox1.SelectedIndex != -1)
                        test.IssueType = comboBox1.Text;
                    else
                        test.IssueType = test.IssueType;
                    //New PC form commpletion
                    if (test.Notes != textBox3.Text && textBox3.Text != "")
                        test.Notes += DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt ") + username.Text + ": " + textBox3.Text + "\r\n";

                    //status update push
                    if (checkBox1.Checked)
                    {
                        test.Status = 'O';
                        test.Assigned = System.DirectoryServices.AccountManagement.UserPrincipal.Current.DisplayName;
                        dbContext.SubmitChanges();
                        this.Close();
                        //Priority Update Push
                        if (checkBox5.Checked)
                        {
                            test.Priority = 'L';
                            dbContext.SubmitChanges();
                            _owner.LoadData();
                            Close();
                        }
                        else if (checkBox6.Checked)
                        {
                            test.Priority = 'M';
                            dbContext.SubmitChanges();
                            _owner.LoadData();
                            Close();
                        }
                        else if (checkBox7.Checked)
                        {
                            test.Priority = 'H';
                            dbContext.SubmitChanges();
                            _owner.LoadData();
                            Close();
                        }

                        //supported by IT update push
                        if (checkBox8.Checked)
                        {
                            test.Supported = true;
                            dbContext.SubmitChanges();
                            this.Close();
                        }
                        else
                        {
                            test.Supported = false;
                            dbContext.SubmitChanges();
                            this.Close();
                        }
                    }
                    else if (checkBox2.Checked)
                    {
                        test.Status = 'P';
                        dbContext.SubmitChanges();
                        this.Close();
                        //Priority Update Push
                        if (checkBox5.Checked)
                        {
                            test.Priority = 'L';
                            dbContext.SubmitChanges();
                            _owner.LoadData();
                            this.Close();
                        }
                        else if (checkBox6.Checked)
                        {
                            test.Priority = 'M';
                            dbContext.SubmitChanges();
                            _owner.LoadData();
                            this.Close();
                        }
                        else if (checkBox7.Checked)
                        {
                            test.Priority = 'H';
                            dbContext.SubmitChanges();
                            _owner.LoadData();
                            this.Close();
                        }

                        //supported by IT update push
                        if (checkBox8.Checked)
                        {
                            test.Supported = true;
                            dbContext.SubmitChanges();
                            this.Close();
                        }
                        else
                        {
                            test.Supported = false;
                            dbContext.SubmitChanges();
                            this.Close();
                        }
                    }
                    else if (checkBox3.Checked)
                    {
                        test.Status = 'C';
                        if (textBox2.Text == "" || textBox4.Text == "" || textBox5.Text == "")
                        {
                            if (test.Resolver == "")
                            {
                                MessageBox.Show("Please fill in all details to close the ticket", "Status update unsuccessful",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else
                            {
                                MessageBox.Show("Ticket Already Closed", "Status update unsuccessful",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                        }
                        else
                        {
                            test.Dresolve = label10.Text;
                            test.Resolver = username.Text;
                            test.Remail = textBox4.Text;
                            test.Rphone = textBox5.Text;
                            test.Resolution = textBox2.Text;

                            bool isadmin = false;
                            if (test.IssueType == "IT-Update" || test.IssueType == "IT-New PC"
                                || test.IssueType == "IT-Permissions" || test.IssueType == "IT-Destroy"
                                || test.IssueType == "IT-Phone" || test.IssueType == "IT-Security"
                                || test.IssueType == "IT-Tablet" || test.IssueType == "IT-User"
                                || test.IssueType == "IT-Update" || test.IssueType == "IT-Other"
                                || test.IssueType == "IT-ODI")
                                isadmin = true;

                            sendemail(useselected, test.Temail, dbContext, isadmin);
                        }
                    }
                    else if (checkBox4.Checked)
                    {
                        test.Status = 'F';
                        dbContext.SubmitChanges();
                        this.Close();
                        //Priority Update Push
                        if (checkBox5.Checked)
                        {
                            test.Priority = 'L';
                            dbContext.SubmitChanges();
                            _owner.LoadData();
                            this.Close();
                        }
                        else if (checkBox6.Checked)
                        {
                            test.Priority = 'M';
                            dbContext.SubmitChanges();
                            _owner.LoadData();
                            this.Close();
                        }
                        else if (checkBox7.Checked)
                        {
                            test.Priority = 'H';
                            dbContext.SubmitChanges();
                            _owner.LoadData();
                            this.Close();
                        }

                        //supported by IT update push
                        if (checkBox8.Checked)
                        {
                            test.Supported = true;
                            dbContext.SubmitChanges();
                            this.Close();
                        }
                        else
                        {
                            test.Supported = false;
                            dbContext.SubmitChanges();
                            this.Close();
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("No server connection available. Please try again.");
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
                    ClosedTicket.sendemail("Your IT Ticket has been closed.", "Thank you for patiently waiting while we worked to resolve your issue.\r\n\r\nYour resolved ticket ID is: " + ticket + "\r\n\r\nThank you,\r\nYour IT Support Team", ticket.ToString(), eto.ToString());
                    try
                    {
                        dbContext.SubmitChanges();
                        this.Close();
                    }
                    catch
                    {
                        MessageBox.Show("Unable to close ticket. No server connection available");
                    }
                    this.Close();
                }
                catch
                {
                    MessageBox.Show("User email address. not valid. Changes made, no email sent.");
                    dbContext.SubmitChanges();
                    this.Close();
                }
            }
            else
            {
                try
                {
                    try
                    {
                        dbContext.SubmitChanges();
                        this.Close();
                    }
                    catch
                    {
                        MessageBox.Show("Unable to close ticket. No server connection available");
                    }
                    this.Close();
                }
                catch
                {
                    MessageBox.Show("User email address. not valid. Changes made, no email sent.");
                    dbContext.SubmitChanges();
                    this.Close();
                }
            }
        }

        //cancel update
        private void Cancel(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var update = new NewPassword(label2.Text);
            update.Show();
        }
    }
}
