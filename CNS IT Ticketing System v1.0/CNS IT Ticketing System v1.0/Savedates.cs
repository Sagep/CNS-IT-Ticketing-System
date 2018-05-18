using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace CNS_IT_Ticketing_System_v1._0
{
    public partial class Savedates : Form
    {
        public Savedates()
        {
            InitializeComponent();
        }

        //Selected dates to print
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime first = new DateTime();
                first = dateTimePicker1.Value;

                DateTime second = new DateTime();
                second = dateTimePicker2.Value;

                TimeSpan duration = second - first;

                int time = (int)duration.TotalMinutes;
                if (time < 1)
                {
                    MessageBox.Show("Start date MUST be before End date");
                }
                else
                {
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "Spreadsheet|*.csv";
                    save.ShowDialog();

                    TicketsDBDataContext db = new TicketsDBDataContext();
                    var r = from p in db.Ticketers
                            select p;

                    string path = save.FileName.ToString();
                    string delimiter = ",";

                    //string builder
                    StringBuilder sb = new StringBuilder();

                    //header
                    string[] hrow = new string[] { "TicketID", "Ticketor", "Date Ticket Created", "Ticketor Email", "Ticketor Phone number", "Device ID","Supported by CNS IT","Assigned to:", "Issue", "Resolver", "Resolver Email", "Resolver Phone Number", "Resolution", "Date Resolved", "Priority", "Issue Type", "Status", "Additional Notes" };
                    sb.AppendLine(string.Join(delimiter, hrow));

                    foreach (var x in r)
                    {
                        DateTime datedata = DateTime.Parse(x.Dissue.ToString());
                        if (datedata > first && datedata < second)
                        {

                            string issue = "";
                            string resolution = "";
                            string notes = "";
                            string ticketoremail = "";
                            string ticketorphone = "";
                            string resolverphone = "";
                            string resolveremail = "";
                            if (x.Temail != null)
                            {
                                ticketoremail = Regex.Replace(x.Temail, @"\r\n?|\n|\t|,", String.Empty);
                            }
                            if (x.Tphone != null)
                            {
                                ticketorphone = Regex.Replace(x.Tphone, @"\r\n?|\n|\t|,", String.Empty);
                            }
                            if (x.Rphone != null)
                            {
                                resolverphone = Regex.Replace(x.Rphone, @"\r\n?|\n|\t|,", String.Empty);
                            }
                            if (x.Remail != null)
                            {
                                resolveremail = Regex.Replace(x.Remail, @"\r\n?|\n|\t|,", String.Empty);
                            }
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
                            string[] row = new string[] { x.ID2.ToString(), x.Ticketor, x.Dissue, ticketoremail, ticketorphone, x.DeviceID, x.Supported.ToString(), x.Assigned, issue, x.Resolver, resolveremail, resolverphone, resolution, x.Dresolve, x.Priority.ToString(), x.IssueType, x.Status.ToString(), notes };
                            sb.AppendLine(string.Join(delimiter, row));
                        }
                        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
                        this.Close();
                    }
                }
            }
            catch
            {
                this.Close();
            }
        }

        private void Savedates_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Now.Date;
            dateTimePicker2.Value = DateTime.Now.Date;

            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "MM-dd-yyyy h:mm tt";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "MM-dd-yyyy h:mm tt";
        }

        //cancel date save
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
