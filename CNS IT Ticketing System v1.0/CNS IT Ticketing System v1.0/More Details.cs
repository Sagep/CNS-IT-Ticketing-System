using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace CNS_IT_Ticketing_System_v1._0
{
    public partial class More_Details : Form
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
                if (curChar+1 < _text.Length)
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

        long CaseID;
        private string DID;

        public More_Details(long ID)
        {
            InitializeComponent();
            CaseID = ID;
        }

        //returns nums amount of character spaced as a string.
        private string spacer(int nums, char spaced)
        {
            string chars = "";
            for (int i=0; i<nums; i++)
            {
                chars += spaced;
            }
            return chars;
        }

        //page load
        private void More_Details_Load(object sender, EventArgs e)
        {
            ID.Text += "" + CaseID;
            label23.Visible = false;
            label22.Visible = false;
            DID = "";
            button3.Visible = false;
            TicketsDBDataContext db = new TicketsDBDataContext();

            var r = from p in db.Ticketers
                    where p.ID2 == CaseID
                    select p;

            foreach (var x in r)
            {
                label2.Text = x.Ticketor;
                if (x.Temail == null)
                {
                    label15.Text = "None";
                }
                else
                    label15.Text = x.Temail;

                if (x.Tphone == null)
                {
                    label17.Text = "None";
                }
                else
                    label17.Text = x.Tphone;
                label4.Text = x.Dissue;
                if(x.Priority=='H')
                {
                    label21.Text = "High";
                }
                else if (x.Priority == 'M')
                {
                    label21.Text = "Medium";
                }
                else if (x.Priority == 'L'||x.Priority=='P')
                {
                    label21.Text = "Low";
                }

                if (x.Status == 'C')
                {
                    label10.Text = x.Resolver;
                    label19.Text = x.Remail;
                    label8.Text = x.Rphone;
                    label12.Text = x.Dresolve;
                }
                else
                {
                    label9.Text = "Current status:";
                    if(x.Status=='O')
                    {
                        label10.Text = "Open";
                    }
                    if (x.Status == 'P')
                    {
                        label10.Text = "Pending";
                    }
                    if (x.Status == 'C')
                    {
                        label10.Text = "Closed";
                    }
                    if (x.Status == 'F')
                    {
                        label10.Text = "Follow-Up Needed";
                    }
                    label18.Visible = false;
                    label7.Visible = false;
                    label11.Visible = false;
                    label19.Visible = false;
                    label8.Visible = false;
                    label12.Visible = false;
                    label22.Visible = false;
                    label23.Visible = false;
                    if (x.DeviceID != "" && x.DeviceID!=null)
                    {
                        label6.Visible = false;
                        textBox2.Visible = false;
                        button3.Visible = true;
                        label22.Visible = true;
                        label23.Visible = true;
                        DID = x.DeviceID;
                        label23.Text = x.DeviceID.ToString();
                        
                    }
                    else
                    {
                        label22.Visible = false;
                        label23.Visible = false;
                    }
                }
                if (x.Supported)
                {
                    label24.Text = "Supported: Yes";
                }
                else
                    label24.Text = "Supported: No";
                textBox1.Text = x.Issue;
                textBox2.Text = x.Resolution;
                textBox3.Text = x.Notes;
            }
        }
        //Shortcut commands
        private void Copy_Click(object sender, EventArgs e)
        {
            Clipboard.SetText("" + CaseID);
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.S))
            {
                EventArgs e = new EventArgs();
                button2_Click(this, e);
                return true;
            }
            else if(keyData == (Keys.Control | Keys.P))
            {
                EventArgs e = new EventArgs();
                button1_Click(this, e);
                return true;
            }
            else if(keyData == (Keys.Escape))
            {
                Close();
                return true; 
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #region buttons
            //print
            private void button1_Click(object sender, EventArgs e)
            {
                //data holders
                string creator = "";
                string cemail = "";
                string cphone = "";
                string DC = "";
                string stat = "";
                string priority = "";
                string Device = "";
                string type = "";
                string resolver = "";
                string daresolv = "";
                string eresolv = "";
                string presolv = "";
                string issue = "";
                string notes = "";
                string resolution = "";
                string supported = "";
                string Assigned = "";

                //load Data
                TicketsDBDataContext db = new TicketsDBDataContext();

                var r = from p in db.Ticketers
                        where p.ID2 == CaseID
                        select p;
                foreach (var x in r)
                {
                    creator = x.Ticketor;
                    cemail = x.Temail;
                    cphone = x.Tphone;
                    DC = x.Dissue;
                    stat = "" + x.Status;
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
                    if (x.Supported)
                        supported = "Issue is Supported by CNS IT";
                    else if (!x.Supported)
                        supported = "Issue is NOT supported by CNS IT";
                    Device = x.DeviceID;
                    type = x.IssueType;
                    resolver = x.Resolver;
                    daresolv = x.Dresolve;
                    eresolv = x.Remail;
                    presolv = x.Rphone;
                    issue = x.Issue;
                    notes = x.Notes;
                    resolution = x.Resolution;
                    Assigned = x.Assigned;
                }

                //get a new class for printing
                PCPrint printer = new PCPrint();
                //sent font
                printer.PrinterFont = new Font("Arial", 10);
                //set text
                printer.TextToPrint = " " + "CNS" + spacer(100, ' ') + "Ticket ID: " + CaseID + "\r\nIT Ticketing System" + "\r\n\r\n\r\n";
                printer.TextToPrint += "Creator of ticket: " + creator + "\r\n";
                printer.TextToPrint += "Creator's email: " + cemail + "\r\n";
                printer.TextToPrint += "Creator's phone Number: " + cphone + "\r\n";
                printer.TextToPrint += "Date Created: " + DC + "\r\n";
                printer.TextToPrint += "Status: " + stat + "\r\n";
                printer.TextToPrint += supported + "\r\n";
                printer.TextToPrint += "Priority: " + priority + "\r\n";
                printer.TextToPrint += "Device ID(If any): " + Device + "\r\n";
                printer.TextToPrint += "Issue Type: " + type + "\r\n";
                printer.TextToPrint += "Assigned to: " + Assigned + "\r\n";
                printer.TextToPrint += "Resolver: " + resolver + "\r\n";
                printer.TextToPrint += "Resolver email: " + eresolv + "\r\n";
                printer.TextToPrint += "Resolver phone Number: " + presolv + "\r\n";
                printer.TextToPrint += "Date Resolved: " + daresolv + "\r\n";
                printer.TextToPrint += "\r\n\r\n" + spacer(70, ' ') + "Issue:\r\n" + spacer(125, '-') + "\r\n" + issue + "\r\n" + spacer(125, '-') + "\r\n\r\n";
                printer.TextToPrint += spacer(65, ' ') + "Additional Notes:\r\n" + spacer(125, '-') + "\r\n" + notes + "\r\n" + spacer(125, '-') + "\r\n\r\n";
                printer.TextToPrint += spacer(67, ' ') + "Resolution\r\n" + spacer(125, '-') + "\r\n" + resolution;
                //print
                printer.Print();
            }

            //save as
            private void button2_Click(object sender, EventArgs e)
            {
                try
                {
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "Spreadsheet|*.csv";
                    save.ShowDialog();

                    string path = save.FileName.ToString();
                    string delimiter = ",";

                    //string builder
                    StringBuilder sb = new StringBuilder();


                    //header
                    string[] hrow = new string[] { "TicketID", "Ticketor", "Date Ticket Created", "Ticketor Email", "Ticketor Phone number", "Device ID", "Supported by CNS IT:", "Assigned To:", "Issue", "Resolver", "Resolver Email", "Resolver Phone Number", "Resolution", "Date Resolved", "Priority", "Issue Type", "Status", "Additional Notes" };
                    sb.AppendLine(string.Join(delimiter, hrow));

                    TicketsDBDataContext db = new TicketsDBDataContext();
                    var r = from p in db.Ticketers
                            where p.ID2 == CaseID
                            select p;

                    foreach (var x in r)
                    {
                        string issue = "";
                        string resolution = "";
                        string notes = "";
                        if (x.Issue != null)
                        {
                            issue = Regex.Replace(x.Issue, @"\r\n?|\n|\t|,", String.Empty);
                        }
                        if (x.Resolution != null)
                        {
                            resolution = String.Join("", x.Resolution.Where(c => c != '\n' && c != '\r' && c != '\t' && c != ','));
                        }
                        if (x.Notes != null)
                            notes = String.Join("", x.Notes.Where(c => c != '\n' && c != '\r' && c != '\t' && c != ','));

                        //new row
                        string[] row = new string[] { x.ID2.ToString(), x.Ticketor, x.Dissue, x.Temail, x.Tphone, x.DeviceID, x.Supported.ToString(), x.Assigned, issue, x.Resolver, x.Remail, x.Rphone, resolution, x.Dresolve, x.Priority.ToString(), x.IssueType, x.Status.ToString(), notes };
                        sb.AppendLine(string.Join(delimiter, row));
                    }
                    //saves rows to file
                    File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
                }
                catch
                { }
            }

            //remote desktop launch
            private void button3_Click(object sender, EventArgs e)
            {
                System.Diagnostics.Process.Start("mstsc.exe", "-v:" + DID);
            }
        #endregion
    }
}
