using System;
using System.Windows.Forms;

namespace CNS_IT_Ticketing_System_v1._0
{
    public partial class AdminTickets : Form
    {
        Form1 _owner;
        public AdminTickets(Form1 owner)
        {
            InitializeComponent();
            _owner = owner;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form2_FormClosing);
        }
        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            _owner.LoadData();
        }
        private void AdminTickets_Load(object sender, EventArgs e)
        {
            textBox2.Text = "itsupport@cnscares.com";
            textBox3.Text = "970-254-9001";
            this.ActiveControl = textBox2;
            label10.Text = "";
            this.textBox1.Hide();
            this.textBox4.Hide();
            this.textBox5.Hide();
            this.textBox6.Hide();
            this.textBox7.Hide();
            this.label15.Visible = false;
            this.label11.Visible = false;
            this.label12.Visible = false;
            this.label13.Visible = false;

            username.Text = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            label5.Text = DateTime.Now.ToString("MM-dd-yyyy h:mm tt");
            label2.Text = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            label4.Text = DateTime.Now.ToString("MM-dd-yyyy h:mm tt");
            Timer timer = new Timer();
            timer.Interval = (900 * 1000);
            timer.Tick += new EventHandler(timer_tick);
            timer.Start();
        }
        private void timer_tick(object sender, EventArgs e)
        {
            this.Close();
        }

        //form close
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

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

                label8.Text = "Computer Setup";
                label10.Text = "Type of computer:";
                label11.Text = "Serial Number of device:";
                label12.Text = "Device CNS-ID: ";
                label13.Text = "Name of Assigne: ";
                textBox1.Text = "";
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                label8.Text = "Change a users Permissions";
                label10.Text = "User Needing changed:";
                label11.Text = "Changed Permission(s): ";
                this.textBox1.Show();
                this.textBox4.Show();
                this.textBox5.Show();
                textBox6.Hide();
                textBox7.Hide();
                this.label11.Visible = true;
                this.label12.Visible = false;
                label13.Visible = false;
                this.label15.Visible = true;
                this.label17.Visible = false;
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                this.textBox1.Show();
                this.textBox4.Show();
                this.textBox5.Show();
                this.textBox6.Show();
                this.textBox7.Hide();
                this.label11.Visible = true;
                this.label12.Visible = true;
                label13.Visible = false;
                this.label15.Visible = true;
                this.label17.Visible = false;


                label8.Text = "Device Destruction";
                label10.Text = "Type of Device:";
                label11.Text = "CNS-ID of device: ";
                label12.Text = "Name of Assigne: ";
                textBox1.Text = "";
            }
            else if (comboBox1.SelectedIndex == 3)
            {
                this.textBox1.Show();
                this.textBox4.Show();
                this.textBox5.Show();
                this.textBox6.Show();
                this.textBox7.Hide();
                this.label11.Visible = true;
                this.label12.Visible = true;
                label13.Visible = false;
                this.label15.Visible = true;
                this.label17.Visible = false;


                label8.Text = "Phone Setup";
                label10.Text = "Type of Device:";
                label11.Text = "CNS-ID of device:";
                label12.Text = "Name of Assigne: ";
                textBox1.Text = "";
            }
            else if (comboBox1.SelectedIndex == 4)
            {
                this.textBox1.Show();
                this.textBox4.Show();
                this.textBox5.Show();
                this.textBox6.Hide();
                this.textBox7.Hide();
                this.label11.Visible = true;
                this.label12.Visible = false;
                label13.Visible = false;
                this.label15.Visible = true;
                this.label17.Visible = false;


                label8.Text = "Security Log Pull";
                label10.Text = "Type of Device:";
                label11.Text = "CNS-ID of device:";
                textBox1.Text = "";
            }
            else if (comboBox1.SelectedIndex == 5)
            {
                this.textBox1.Show();
                this.textBox4.Show();
                this.textBox5.Show();
                this.textBox6.Hide();
                this.textBox7.Hide();
                this.label11.Visible = true;
                this.label12.Visible = false;
                label13.Visible = false;
                this.label15.Visible = true;
                this.label17.Visible = false;


                label8.Text = "Tablet Reconfigure-Setup";
                label10.Text = "Reconfiguration or Setup?";
                label11.Text = "CNS-ID of device:";
                textBox1.Text = "";
            }
            else if (comboBox1.SelectedIndex == 6)
            {
                this.textBox1.Show();
                this.textBox4.Show();
                this.textBox5.Show();
                this.textBox6.Show();
                this.textBox7.Show();
                this.label11.Visible = true;
                this.label12.Visible = true;
                label13.Visible = true;
                this.label15.Visible = true;
                this.label17.Visible = false;


                label8.Text = "User Setup";
                label10.Text = "Agent title: ";
                label11.Text = "Date agent Starts: ";
                label12.Text = "Drives Agent needs: ";
                label13.Text = "Agent Name";
                textBox1.Text = "";
            }
            else if (comboBox1.SelectedIndex == 7)
            {
                this.textBox1.Show();
                this.textBox4.Hide();
                this.textBox5.Hide();
                this.textBox6.Hide();
                this.textBox7.Hide();
                this.label11.Visible = false;
                this.label12.Visible = false;
                label13.Visible = false;
                this.label15.Visible = true;
                this.label17.Visible = false;


                label8.Text = "Update Logs Pulled";
                label10.Text = "";
                textBox1.Text = "";
            }
            else if (comboBox1.SelectedIndex == 8)
            {
                this.textBox1.Show();
                this.textBox4.Hide();
                this.textBox5.Hide();
                this.textBox6.Hide();
                this.textBox7.Hide();
                this.label11.Visible = false;
                this.label12.Visible = false;
                label13.Visible = false;
                this.label15.Visible = true;
                this.label17.Visible = false;


                label8.Text = "Other";
                label10.Text = "";
                textBox1.Text = "";
            }
            else if (comboBox1.SelectedIndex == 9)
            {
                this.textBox1.Show();
                this.textBox4.Hide();
                this.textBox5.Hide();
                this.textBox6.Hide();
                this.textBox7.Hide();
                this.label11.Visible = false;
                this.label12.Visible = false;
                label13.Visible = false;
                this.label15.Visible = true;
                this.label17.Visible = false;


                label8.Text = "Off Duty Issue";
                label10.Text = "";
                textBox1.Text = "";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("User MUST select a ticket type", "Error, invalid selection", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Random rnd = new Random();
                long ticketnumber = long.Parse(DateTime.Now.ToString("MMddyyyyhhmm") + rnd.Next(1, 1000).ToString());
                string ticketor = "";
                string tphone = "";
                string temail = "";
                string dissue = "";
                string issue = "";
                char status = 'P';
                string Device = "";
                string issuetype = label8.Text;
                ticketor = username.Text;
                tphone = textBox3.Text;
                temail = textBox2.Text;
                dissue = label5.Text;

                if (comboBox1.SelectedIndex == 0)
                {
                    issuetype = "IT-New PC";
                    issue = "New Computer Setup. \r\nType of Device: " + textBox4.Text + "\r\nDevice-ID: " + textBox6.Text + "\r\nSerial Number of device: " + textBox5.Text + "\r\nDevice is Assigned to: " + textBox7.Text + "\r\n\r\nMore Details:\r\n" + textBox1.Text;
                    Device = textBox6.Text;
                }
                else if (comboBox1.SelectedIndex ==1)
                {
                    issuetype = "IT-Permissions";
                    issue = "Change of user permissions\r\nName of agent needing changed: " + textBox4.Text + "\r\nWhat needs changed: " + textBox5.Text;
                }
                else if (comboBox1.SelectedIndex == 2)
                {
                    issuetype = "IT-Destroy";
                    issue = "Device Destruction. \r\nType of Device: \r\n"+textBox4.Text+"\r\nCNS-ID of Device: "+textBox5.Text+"\r\nPrevious person it was assigned to: "+textBox6.Text+"\r\n\r\nMore Details:\r\n"+textBox1.Text;
                    Device = textBox5.Text;
                }
                else if (comboBox1.SelectedIndex == 3)
                {
                    issuetype = "IT-Phone";
                    issue = "Phone Setup\r\nType of Device: " + textBox4.Text + "\r\nCNS-ID of device:" + textBox5.Text + "\r\nName of person it is being assigned to: " + textBox6.Text + "\r\n\r\nMore Details:\r\n" + textBox1.Text;
                    Device = textBox5.Text;
                }
                else if (comboBox1.SelectedIndex == 4)
                {
                    issuetype = "IT-Security";
                    issue = "Security Log Pull\r\nType of Device: " + textBox4.Text + "\r\nCNS-ID of device: " + textBox5.Text + "\r\n\r\nMore Details:\r\n" + textBox1.Text;
                    Device = textBox5.Text;
                }
                else if (comboBox1.SelectedIndex == 5)
                {
                    issuetype = "IT-Tablet";
                    issue = "Tablet Setup\r\nIs it being reconfigured or setup: " + textBox4.Text + "\r\nCNS-ID of device: " + textBox5.Text + "\r\n\r\nMore Details:\r\n" + textBox1.Text;
                }
                else if(comboBox1.SelectedIndex == 6)
                {
                    issuetype = "IT-User";
                    issue = "New user Setup\r\nAgent Name: " + textBox7.Text + "\r\nAgent Title: " + textBox4.Text + "\r\nDate Agent Starts: " + textBox5.Text + "\r\nDrives Agent needs: " + textBox6.Text + "\r\n\r\nMore Details:\r\n" + textBox1.Text;
                }
                else if (comboBox1.SelectedIndex ==7)
                {
                    issuetype = "IT-Update";
                    issue = "Update Logs Pulled from N-able" + "\r\n\r\nMore Details:\r\n" + textBox1.Text;
                }
                else if (comboBox1.SelectedIndex == 8)
                {
                    issuetype = "IT-Other";
                    issue = "Other" + "\r\n\r\nMore Details:\r\n" + textBox1.Text;
                }
                else if (comboBox1.SelectedIndex == 9)
                {
                    issuetype = "IT-ODI";
                    issue = "Off Duty Issue" + "\r\n\r\nMore Details:\r\n" + textBox1.Text;

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
                        IssueType = issuetype,
                        Resolver = null,
                        Priority = 'P',
                        Supported = true
                    };
                    dbContext.Ticketers.InsertOnSubmit(test);
                    try
                    {
                        dbContext.SubmitChanges();
                        MessageBox.Show("Your ticket is submited. The ID is: " + ticketnumber, "Ticket Creation Successful!",
MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch
                    {
                        MessageBox.Show("Unable to create a new ticket. No server connection available");
                    }
                    this.Close();
                }


            }
        }
    }
}
