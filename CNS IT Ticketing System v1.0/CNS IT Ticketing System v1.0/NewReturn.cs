using System;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;

namespace CNS_IT_Ticketing_System_v1._0
{
    public partial class NewReturn : Form
    {
        string username;
        Form1 _owner;
        public NewReturn(string user, Form1 owner)
        {
            InitializeComponent();
            _owner = owner;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.NewReturn_FormClosing);
            username = user;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void NewReturn_FormClosing(object sender, FormClosingEventArgs e)
        {
            _owner.LoadData();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if ((comboBox1.SelectedIndex == 4 && textBox2.Text == "")||textBox1.Text=="" || textBox3.Text=="")
            {
                MessageBox.Show("Please enter details in all fields");
            }
            else
            {
                using (EquipReDataContext dbContext = new EquipReDataContext())
                {
                    DateTime blank = new DateTime(1990, 1, 1, 0, 0, 0);

                    Random rnd = new Random();
                    string ID = "E" + long.Parse(DateTime.Now.ToString("MMddyyyyhhmm") + rnd.Next(1, 1000).ToString());

                    string reasons = "";
                    if (textBox2.Text != "")
                    {
                        reasons = comboBox1.Text + ": " + textBox2.Text;
                    }
                    else
                        reasons = comboBox1.Text;
                    EquipmentReturn returns = new EquipmentReturn()
                    {
                        Id = long.Parse(DateTime.Now.ToString("MMddyyyyhhmm") + rnd.Next(1, 1000).ToString()),
                        IDText = ID,
                        FirstContactPerson = "None",
                        SecondContactPerson = "None",
                        ThirdContactperson = "None",
                        EscalationAgent = "None",
                        EscalationStatus = false,

                        FirstContactDate = blank,
                        SecondContactDate = blank,
                        ThirdContactDate = blank,
                        EscalationDate = blank,
                        Notes = "Device Assigned to: "+textBox3.Text+"\r\n",
                        TicketStatus = 'P',
                        Priority = 'L',
                        Phone_Number = "000-000-0000",

                        Assignee = textBox3.Text,

                        Ticketor = username,
                        DateCreated = DateTime.Now,
                        Reason = reasons,
                        TagID = textBox1.Text
                    };
                    dbContext.EquipmentReturns.InsertOnSubmit(returns);
                    try
                    {
                        dbContext.SubmitChanges();
                    }
                    catch (Exception r)
                    {
                        MessageBox.Show("Error, unable to save to database. Check connection\r\n\r\n" + r.ToString());
                    }
                    email ReturnSetup = new email();
                    ReturnSetup.sendemail("Return Ticket Created.", "A new Return ticket has been created.Please review this in the IT ticketing system for further information.\r\nThe ticket ID is: " + ID + "\r\nEquipmentID: " + textBox1.Text + "Assigned to: " + textBox3.Text, ID, "EquipmentReturns@cnscares.com");

                    this.Close();
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 4)
            {
                textBox2.Show();
                label3.Show();
            }
            else
            {
                textBox2.Hide();
                label3.Hide();
            }
        }

        private void NewReturn_Load(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox2.Hide();
            label3.Hide();
            comboBox1.SelectedIndex = 0;
        }
    }
}
