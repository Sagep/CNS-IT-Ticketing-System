using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;

namespace CNS_IT_Ticketing_System_v1._0
{

    public partial class ReturnData : Form
    {
        public class PCPrint : System.Drawing.Printing.PrintDocument
        {
            #region  Property Variables 
            /// <summary>
            /// Property variable for the Font the user wishes to use
            /// </summary>
            /// <remarks></remarks>
            private Font _font;

            /// <summary>
            /// Property variable for the text to be printed
            /// </summary>
            /// <remarks></remarks>
            private string _text;
            #endregion

            #region  Class Properties 
            /// <summary>
            /// Property to hold the text that is to be printed
            /// </summary>
            /// <value></value>
            /// <returns>A string</returns>
            /// <remarks></remarks>
            public string TextToPrint
            {
                get { return _text; }
                set { _text = value; }
            }

            /// <summary>
            /// Property to hold the font the users wishes to use
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public Font PrinterFont
            {
                // Allows the user to override the default font
                get { return _font; }
                set { _font = value; }
            }
            #endregion

            #region Static Local Variables
            /// <summary>
            /// Static variable to hold the current character
            /// we're currently dealing with.
            /// </summary>
            static int curChar;
            #endregion

            #region  Class Constructors 
            /// <summary>
            /// Empty constructor
            /// </summary>
            /// <remarks></remarks>
            public PCPrint() : base()
            {
                //Set the file stream
                //Instantiate out Text property to an empty string
                _text = string.Empty;
            }

            /// <summary>
            /// Constructor to initialize our printing object
            /// and the text it's supposed to be printing
            /// </summary>
            /// <param name=str>Text that will be printed</param>
            /// <remarks></remarks>
            public PCPrint(string str) : base()
            {
                //Set the file stream
                //Set our Text property value
                _text = str;
            }
            #endregion

            #region  onbeginPrint 
            /// <summary>
            /// Override the default onbeginPrint method of the PrintDocument Object
            /// </summary>
            /// <param name=e></param>
            /// <remarks></remarks>
            protected void onbeginPrint(System.Drawing.Printing.PrintEventArgs e)
            {
                //Check to see if the user provided a font
                //if they didn't then we default to Times New Roman
                if (_font == null)
                {
                    //Create the font we need
                    _font = new Font("Times New Roman", 10);
                }
            }
            #endregion

            #region  OnPrintPage 
            /// <summary>
            /// Override the default OnPrintPage method of the PrintDocument
            /// </summary>
            /// <param name=e></param>
            /// <remarks>This provides the print logic for our document</remarks>
            protected override void OnPrintPage(System.Drawing.Printing.PrintPageEventArgs e)
            {
                // Run base code
                base.OnPrintPage(e);

                //Declare local variables needed

                int printHeight;
                int printWidth;
                int leftMargin;
                int rightMargin;
                Int32 lines;
                Int32 chars;

                //Set print area size and margins
                {
                    printHeight = base.DefaultPageSettings.PaperSize.Height - base.DefaultPageSettings.Margins.Top - base.DefaultPageSettings.Margins.Bottom;
                    printWidth = base.DefaultPageSettings.PaperSize.Width - base.DefaultPageSettings.Margins.Left - base.DefaultPageSettings.Margins.Right;
                    leftMargin = base.DefaultPageSettings.Margins.Left;  //X
                    rightMargin = base.DefaultPageSettings.Margins.Top;  //Y
                }

                //Check if the user selected to print in Landscape mode
                //if they did then we need to swap height/width parameters
                if (base.DefaultPageSettings.Landscape)
                {
                    int tmp;
                    tmp = printHeight;
                    printHeight = printWidth;
                    printWidth = tmp;
                }

                //Now we need to determine the total number of lines
                //we're going to be printing
                Int32 numLines = (int)printHeight / PrinterFont.Height;

                //Create a rectangle printing are for our document
                RectangleF printArea = new RectangleF(leftMargin, rightMargin, printWidth, printHeight);

                //Use the StringFormat class for the text layout of our document
                StringFormat format = new StringFormat(StringFormatFlags.LineLimit);

                //Fit as many characters as we can into the print area      

                e.Graphics.MeasureString(_text.Substring(RemoveZeros(curChar)), PrinterFont, new SizeF(printWidth, printHeight), format, out chars, out lines);

                //Print the page
                e.Graphics.DrawString(_text.Substring(RemoveZeros(curChar)), PrinterFont, Brushes.Black, printArea, format);

                //Increase current char count
                curChar += chars;

                //Detemine if there is more text to print, if
                //there is the tell the printer there is more coming
                if (curChar + 1 < _text.Length)
                {
                    e.HasMorePages = true;
                }
                else
                {
                    e.HasMorePages = false;
                    curChar = 0;
                }
            }

            #endregion

            #region  RemoveZeros 
            /// <summary>
            /// Function to replace any zeros in the size to a 1
            /// Zero's will mess up the printing area
            /// </summary>
            /// <param name=value>Value to check</param>
            /// <returns></returns>
            /// <remarks></remarks>
            public int RemoveZeros(int value)
            {
                //Check the value passed into the function,
                //if the value is a 0 (zero) then return a 1,
                //otherwise return the value passed in
                switch (value)
                {
                    case 0:
                        return 1;
                    default:
                        return value;
                }
            }
            #endregion
        }

        string RID = "";
        string username = "";
        bool clicked = false;
        bool clicked2 = false;
        public ReturnData(string ID, Form1 owner, string user)
        {
            InitializeComponent();
            username = user;
            RID = ID;
        }

        //loading of data
        private void ReturnData_Load(object sender, EventArgs e)
        {
            button1.Focus();
            button2.Enabled = false;
            textBox1.Enabled = true;
            label2.Text = RID;

            EquipReDataContext edb = new EquipReDataContext();
            var r = from p in edb.EquipmentReturns
                    where p.IDText.Contains(RID)
                    select p;

            foreach (var results in r)
            {
                label4.Text = results.Ticketor;
                label6.Text = results.DateCreated.ToString();
                label8.Text = results.TagID;
                label10.Text = results.Reason;
                textBox3.Text = results.AdditionalNotes;

                if (results.Priority == 'H')
                {
                    label25.Text = "High";
                    comboBox1.SelectedIndex = 0;
                }
                else if (results.Priority == 'M')
                {
                    label25.Text = "Medium";
                    comboBox1.SelectedIndex = 1;
                }
                else if (results.Priority == 'L')
                {
                    label25.Text = "Low";
                    comboBox1.SelectedIndex = 2;
                }

                if (results.TicketStatus == 'P')
                {
                    label27.Text = "Pending Resolution";
                    radioButton3.Checked = true;
                }
                else if (results.TicketStatus == 'O')
                {
                    label27.Text = "Open";
                    radioButton1.Checked = true;
                }
                else if (results.TicketStatus == 'F')
                {
                    label27.Text = "Follow Up";
                    radioButton4.Checked = true;
                }
                else if (results.TicketStatus == 'C')
                {
                    radioButton2.Checked = true;
                    label27.Text = "Closed";
                    button1.Enabled = false;
                    button2.Enabled = false;

                    radioButton1.Enabled = false;
                    button5.Enabled = false;
                    radioButton2.Enabled = false;
                    radioButton3.Enabled = false;
                    radioButton4.Enabled = false;
                    comboBox1.Enabled = false;
                    textBox1.ReadOnly = true;
                    textBox2.ReadOnly = true;
                }

                if (results.EscalationStatus == true)
                {
                    label23.Text = "Yes";
                    label29.Text = results.EscalationAgent;
                    label31.Text = results.EscalationDate.ToString();
                    button2.Enabled = false;
                }
                else if (results.EscalationStatus == false)
                {
                    label23.Text = "No";
                    label28.Hide();
                    label29.Hide();
                    label30.Hide();
                    label31.Hide();
                }

                label29.Text = results.EscalationAgent;
                label31.Text = results.EscalationDate.ToString();



                label16.Text = results.FirstContactPerson;
                label17.Text = results.SecondContactPerson;
                label18.Text = results.ThirdContactperson;

                label19.Text = results.FirstContactDate.ToString();
                label20.Text = results.SecondContactDate.ToString();
                label21.Text = results.ThirdContactDate.ToString();

                if (results.FirstContactPerson == "None")
                {
                    label16.Text = "No Contact Made";
                    label14.Hide();
                    label15.Hide();
                    label17.Hide();
                    label18.Hide();
                    label20.Hide();
                    label21.Hide();
                    label19.Text = "";
                }
                else if (results.SecondContactPerson == "None")
                {
                    label17.Text = "No Contact Made";
                    label15.Hide();
                    label18.Hide();
                    label21.Hide();
                    label20.Text = "";
                }
                else if (results.ThirdContactperson == "None")
                {
                    label18.Text = "No Contact Made";
                    label21.Text = "";
                }
                else
                {
                    button1.Enabled = false;
                    button5.Enabled = false;
                    if (results.EscalationStatus == false)
                    {
                        button2.Enabled = true;
                        textBox2.Enabled = false;
                        textBox3.Enabled = false;
                    }
                    else
                    {
                        button2.Enabled = false;
                        textBox2.Enabled = false;
                        textBox3.Enabled = false;
                    }
                }

                if (results.Phone_Number != "000-000-0000")
                {
                    textBox1.Text = results.Phone_Number;
                    textBox1.Enabled = false;
                }

                textBox2.Text = results.Notes;
            }
        }

        //cleartext on click
        private void textBox2_Click(object sender, EventArgs e)
        {
            if (clicked == false && textBox2.ReadOnly != true)
            {
                textBox2.Text = "";
                clicked = true;
            }
        }
        private void textBox3_Click(object sender, EventArgs e)
        {
            if (clicked2 == false && textBox2.ReadOnly != true)
            {
                textBox3.Text = "";
                clicked2 = true;
            }
        }

        //submit a new call 
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
                MessageBox.Show("Error! Enter all details before continuing");
            else
            {
                using (EquipReDataContext dbContext = new EquipReDataContext())
                {
                    try
                    {
                        EquipmentReturn submitchanges = dbContext.EquipmentReturns.SingleOrDefault(X => X.IDText.Contains(RID));

                        //notes
                        if (textBox2.Text != "" && submitchanges.Notes == "")
                        {
                            submitchanges.Notes += username + " (" + DateTime.Now.ToString() + "): " + textBox2.Text;
                        }
                        else if (textBox2.Text != "" && submitchanges.Notes != "" && submitchanges.Notes != textBox2.Text)
                        {
                            submitchanges.Notes += "\r\n\r\n" + username + " (" + DateTime.Now.ToString() + "): " + textBox2.Text;
                        }

                        //additional notes
                        if (textBox3.Text != "" && submitchanges.AdditionalNotes == "")
                        {
                            submitchanges.AdditionalNotes += username + " (" + DateTime.Now.ToString() + "): " + textBox3.Text;
                        }
                        else if (textBox3.Text != "" && submitchanges.AdditionalNotes != "" && submitchanges.AdditionalNotes != textBox2.Text)
                        {
                            submitchanges.AdditionalNotes += "\r\n\r\n" + username + " (" + DateTime.Now.ToString() + "): " + textBox3.Text;
                        }

                        //contacts
                        if (submitchanges.FirstContactPerson == "None")
                        {
                            submitchanges.FirstContactPerson = username;
                            submitchanges.FirstContactDate = DateTime.Now;
                        }
                        else if (submitchanges.SecondContactPerson == "None")
                        {
                            submitchanges.SecondContactPerson = username;
                            submitchanges.SecondContactDate = DateTime.Now;
                        }
                        else if (submitchanges.ThirdContactperson == "None")
                        {
                            submitchanges.ThirdContactperson = username;
                            submitchanges.ThirdContactDate = DateTime.Now;
                        }

                        //phone number
                        if (submitchanges.Phone_Number == "000-000-0000")
                            submitchanges.Phone_Number = textBox1.Text;
                        
                        if (comboBox1.SelectedIndex == 0)
                            submitchanges.Priority = 'H';
                        if (comboBox1.SelectedIndex == 1)
                            submitchanges.Priority = 'M';
                        if (comboBox1.SelectedIndex == 2)
                            submitchanges.Priority = 'L';

                        if (submitchanges.Notes == textBox2.Text || (textBox2.Text == "" && submitchanges.Notes == "") || textBox2.Text == "")
                            MessageBox.Show("Error! Enter all details before continuing");
                        else
                        {
                            //Status
                            if (radioButton1.Checked == true)
                                submitchanges.TicketStatus = 'O';
                            if (radioButton2.Checked == true)
                            {
                                email ReturnClosed = new email();
                                ReturnClosed.sendemail("Return Ticket Closed", "This is a notice that a ticket for equipment returns has been closed. \r\n\r\nDevice ID: " + submitchanges.TagID + "\r\nPhone Number Called: " + submitchanges.Phone_Number + "\r\nDevice Assigned to: " + submitchanges.Assignee, submitchanges.IDText, "EquipmentReturns@cnscares.com");
                                submitchanges.TicketStatus = 'C';
                            }
                            if (radioButton3.Checked == true)
                                submitchanges.TicketStatus = 'P';
                            if (radioButton4.Checked == true)
                                submitchanges.TicketStatus = 'F';
                            dbContext.SubmitChanges();
                            this.Close();
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Error Submiting Data. Please contact your Network Administrator or IT team.");
                    }
                }
            }
        }

        //cancel button
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var escalater = new EscalateReturn(RID, this);
            escalater.Show();
        }


        //print functions
        private string spacer(int nums, char spaced)
        {
            string chars = "";
            for (int i = 0; i < nums; i++)
            {
                chars += spaced;
            }
            return chars;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //data holders
            string creator = "";
            string DC = "";
            string stat = "";
            string priority = "";
            string Device = "";
            string type = "";
            string notes = "";
            string adnotes = "";
            //people who have handled
            string fp = "";
            string sp = "";
            string tp = "";
            string ep = "";
            //dates of when handled
            string fd = "";
            string sd = "";
            string td = "";
            string ed = "";
            //escalation status
            string es = "";
            //reason why returning
            string reason = "";
            //Phone Number
            string phone = "";
            //Assignee name
            string assigned = "";

            //load Data
            EquipReDataContext db = new EquipReDataContext();

            var r = from p in db.EquipmentReturns
                    where p.IDText.Contains(RID)
                    select p;
            foreach (var x in r)
            {
                creator = x.Ticketor;
                DC = x.DateCreated.ToString();
                stat = "" + x.TicketStatus;
                if (stat == "F")
                    stat = "Follow-Up";
                if (stat == "C")
                    stat = "Closed";
                if (stat == "P")
                    stat = "Pending";
                if (stat == "O")
                    stat = "Open";
                priority = x.Priority + "";
                if (priority == "L")
                    priority = "Low";
                if (priority == "M")
                    priority = "Medium";
                if (priority == "H")
                    priority = "High";
                phone = x.Phone_Number;
                Device = x.TagID;
                type = "Return";
                notes = x.Notes;
                adnotes = x.AdditionalNotes;
                //contact people
                fp = x.FirstContactPerson;
                sp = x.SecondContactPerson;
                tp = x.ThirdContactperson;
                ep = x.EscalationAgent;
                assigned = x.Assignee;
                //contact dates
                if (x.FirstContactDate.ToString() != "01/01/1990 12:00:00 AM")
                {
                    fp = x.FirstContactPerson;
                    fd = x.FirstContactDate.ToString();
                }
                else
                {
                    fp = "";
                    fd = "No contact";
                }
                if (x.SecondContactDate.ToString() != "01/01/1990 12:00:00 AM")
                {
                    sp = x.SecondContactPerson;
                    sd = x.SecondContactDate.ToString();
                }
                else
                {
                    sp = "";
                    sd = "No contact";
                }
                if (x.ThirdContactDate.ToString() != "01/01/1990 12:00:00 AM")
                {
                    tp = x.ThirdContactperson;
                    td = x.ThirdContactDate.ToString();
                }
                else
                {
                    tp = "";
                    td = "No contact";
                }
                reason = x.Reason;
                //escalation status
                if (x.EscalationStatus == true)
                {
                    es = "Escalated";
                    ep = x.EscalationAgent;
                    ed = x.EscalationDate.ToString();
                }
                else
                {
                    es = "Not Escalated";
                    ep = "";
                    ed = "No Contact";
                }
            }

            //get a new class for printing
            PCPrint printer = new PCPrint();
            //sent font
            printer.PrinterFont = new Font("Arial", 10);
            //set text
            printer.TextToPrint = " " + "CNS" + spacer(100, ' ') + "Ticket ID: " + RID + "\r\nIT Ticketing System" + "\r\n\r\n\r\n";
            printer.TextToPrint += "Creator of ticket: " + creator + "\r\n";
            printer.TextToPrint += "Date Created: " + DC + "\r\n";
            printer.TextToPrint += "Status: " + stat + "\r\n";
            printer.TextToPrint += "Priority: " + priority + "\r\n";
            printer.TextToPrint += "Device ID: " + Device + "\r\n";
            printer.TextToPrint += "Issue Type: " + type + "\r\n";
            printer.TextToPrint += "Reason: " + reason + "\r\n";
            printer.TextToPrint += "Device Assigned to: " + assigned + "\r\n";
            printer.TextToPrint += "Phone Number Called: " + phone + "\r\n";
            printer.TextToPrint += "First contact: " + fp + " (" + fd + ")\r\n";
            printer.TextToPrint += "Second contact: " + sp + " (" + sd + ")\r\n";
            printer.TextToPrint += "Third contact: " + tp + " (" + td + ")\r\n";
            printer.TextToPrint += es + "\r\n";
            printer.TextToPrint += "Escalation information: " + ep + " (" + ed + ")\r\n\r\n";
            printer.TextToPrint += spacer(65, ' ') + "Notes:\r\n" + spacer(125, '-') + "\r\n" + notes + "\r\n" + spacer(125, '-') + "\r\n\r\n";
            printer.TextToPrint += spacer(65, ' ') + "Additional Notes:\r\n" + spacer(125, '-') + "\r\n" + adnotes + "\r\n" + spacer(125, '-') + "\r\n\r\n";
            //print
            printer.Print();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (EquipReDataContext dbContext = new EquipReDataContext())
            {
                try
                {
                    EquipmentReturn submitchanges = dbContext.EquipmentReturns.SingleOrDefault(X => X.IDText.Contains(RID));

                    //notes
                    if (textBox3.Text != "" && submitchanges.AdditionalNotes == "")
                    {
                        submitchanges.AdditionalNotes += username + " (" + DateTime.Now.ToString() + "): " + textBox3.Text;
                    }
                    else if (textBox3.Text != "" && submitchanges.AdditionalNotes != "" && submitchanges.AdditionalNotes != textBox2.Text)
                    {
                        submitchanges.AdditionalNotes += "\r\n\r\n" + username + " (" + DateTime.Now.ToString() + "): " + textBox3.Text;
                    }


                        submitchanges.TicketStatus = 'F';

                    //level of priority
                    if (comboBox1.SelectedIndex == 0)
                        submitchanges.Priority = 'H';
                    if (comboBox1.SelectedIndex == 1)
                        submitchanges.Priority = 'M';
                    if (comboBox1.SelectedIndex == 2)
                        submitchanges.Priority = 'L';

                    if (submitchanges.AdditionalNotes == textBox3.Text || (textBox3.Text == "" && submitchanges.AdditionalNotes == "") || textBox3.Text == "")
                        MessageBox.Show("Error! Enter all details before continuing");
                    else
                    {
                        //Status
                        if (radioButton1.Checked == true)
                            submitchanges.TicketStatus = 'O';
                        if (radioButton2.Checked == true)
                        {                           
                            email ReturnClosed = new email();
                            ReturnClosed.sendemail("Return Ticket Closed", "This is a notice that a ticket for equipment returns has been closed. \r\n\r\nDevice ID: " + submitchanges.TagID + "\r\nPhone Number Called: " + submitchanges.Phone_Number + "\r\nDevice Assigned to: " + submitchanges.Assignee, submitchanges.IDText, "EquipmentReturns@cnscares.com");
                            submitchanges.TicketStatus = 'C';
                        }
                        if (radioButton3.Checked == true)
                            submitchanges.TicketStatus = 'P';
                        if (radioButton4.Checked == true)
                        {
                            submitchanges.TicketStatus = 'F';
                        }

                        dbContext.SubmitChanges();
                        this.Close();
                    }
                }
                catch
                {
                    MessageBox.Show("Error Submiting Data. Please contact your Network Administrator or IT team.");
                }
            }
        }

        private void sendemail(string TagID, string eto, String phonenumber, string NameOfAssignee)
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
                        msg.Subject = "Return Ticket Closed";
                        StringBuilder sb = new StringBuilder();
                        sb.AppendLine("This is a notice that a ticket for equipment returns has been closed. \r\n\r\nDevice ID: " + TagID + "\r\nPhone Number Called: " + phonenumber + "\r\nDevice Assigned to: " + NameOfAssignee + "\r\n\r\n\r\n");
                        sb.AppendLine("\r\n\r\nPlease note: This is an automated Message. Any replies to this email will be forwarded to itsupport@cnscares.com");
                        msg.Body = sb.ToString();
                        msg.IsBodyHtml = false;
                        smtp.Send(msg);
                    }
                }
                try
                {
                    MessageBox.Show("Ticket Has been Closed", "Ticket closure Successful!",
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

            }
        }
    }
}
