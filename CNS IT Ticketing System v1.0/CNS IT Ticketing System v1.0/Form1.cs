using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Collections;
//checking login privelidges
using System.DirectoryServices.AccountManagement;
//flashing windows for admins
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Text;
using System.IO;

namespace CNS_IT_Ticketing_System_v1._0
{
    //Check if domain admin
    public partial class Form1 : Form
    {
        double version = 4.51;
        bool DomainAdmin;
        bool D365user;
        bool ReturnUser;

        //ticket counters
        int counter;
        bool dcount;
        bool adcount;
        bool ncount;
        bool ccount;
        int scounter;
        Timer timer = new Timer();

        public Form1()
        {
            InitializeComponent();
            this.listView1.View = View.Details;
            //counter initialize
            counter = 0;
            scounter = 0;
            dcount=false;
            adcount = false;
            ncount = false;
            ccount = false;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            timer.Stop();
            //gets the datacontext.
            ChaterDataContext messages = new ChaterDataContext();

            //checks who is logged in. Sets them to be signed out on closure.
            InternalChat test = messages.InternalChats.SingleOrDefault(X => X.Id == 1);
            if (System.Security.Principal.WindowsIdentity.GetCurrent().Name.Contains("Sage"))
            {
                test.SageConnection = false;
            }
            if (System.Security.Principal.WindowsIdentity.GetCurrent().Name.Contains("Mike"))
            {
                test.MikeConnection = false;
            }
            messages.SubmitChanges();
        }

        protected void Displaynotify()
        {
            try
            {
                notifyIcon1.BalloonTipTitle = "IT Ticketing System";
                notifyIcon1.BalloonTipText = "A New ticket has been created";
                notifyIcon1.ShowBalloonTip(100);
            }
            catch (Exception ex)
            {
            }
        }
        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                //Hide();
            }
        }

        #region DataInListview
        //Initial load-up of the program
        private void Form1_Load(object sender, EventArgs e)
        {
            notifyIcon1.Icon = Properties.Resources.favicon2;
            notifyIcon1.Text = "CNS IT Ticketing System";
            notifyIcon1.Visible = true;

            newReturnToolStripMenuItem.Visible = false;
            newAdminToolStripMenuItem.Visible = false;
            updateTicketToolStripMenuItem.Visible = false;
            deleteTicketToolStripMenuItem.Visible = false;
            saveTicketsToolStripMenuItem.Visible = false;
            closeTicketToolStripMenuItem.Visible = false;
            launchChatToolStripMenuItem.Visible = false;
            changeTabToolStripMenuItem.Visible = false;
            saveByTicketToolStripMenuItem.Visible = false;

            //DateTime mountain = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.FindSystemTimeZoneById("Mountain Standard Time"));

            //MessageBox.Show("Current Time: " + DateTime.Now.ToString("MMddyyyy hh:mm")+"\r\n"+ mountain.ToString("MMddyyyy hh:mm"));



            //buttons hide
            radioButton1.Visible = false;
            radioButton2.Visible = false;
            radioButton3.Visible = false;
            radioButton4.Visible = false;
            Returns.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            radioButton1.Select();
            // Create the ToolTip and associate with the Form container.
            ToolTip toolTip1 = new ToolTip();

            // Set up the delays for the ToolTip.
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 1000;
            toolTip1.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip1.ShowAlways = true;

            toolTip1.SetToolTip(button3, "CTRL+N");
            //toolTip1.SetToolTip(save, "CTRL+S");
            toolTip1.SetToolTip(button5, "CTRL+U");
            toolTip1.SetToolTip(button1, "CTRL+D");
            //toolTip1.SetToolTip(button8, "CTRL+H");



            try
            {
                textBox1.ContextMenu = new ContextMenu();
                this.ActiveControl = button3;
                listView1.View = View.Details;
                listView1.FullRowSelect = true;


                timer.Interval = (900 * 1000);
                //timer.Interval = (20 * 1000);
                timer.Tick += new EventHandler(timer_tick);
                timer.Start();
                DomainAdmin = false;

                username.Text = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                DateTime mountain = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.FindSystemTimeZoneById("Mountain Standard Time"));
                label3.Text = mountain.ToString("MM-dd-yyyy h:mm tt")+" (MST)";

                // set up domain context
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "192.168.10.5");

                // find a user
                UserPrincipal user = UserPrincipal.FindByIdentity(ctx, username.Text.ToString());

                // find the group in question
                GroupPrincipal group = GroupPrincipal.FindByIdentity(ctx, "Domain Admins");
                GroupPrincipal group2 = GroupPrincipal.FindByIdentity(ctx, "D365Supports");
                GroupPrincipal group3 = GroupPrincipal.FindByIdentity(ctx, "EquipmentReturns");
                D365user = user.IsMemberOf(group2);
                DomainAdmin = user.IsMemberOf(group);
                ReturnUser = user.IsMemberOf(group3);

                button1.Visible = false;
                button2.Visible = false;
                button5.Visible = false;
                Delete.Visible = false;
                Admin.Visible = false;
                //save.Visible = false;
                label5.Text = "Domain User";

                ////this code tests the access for each user based on the data they are allowed to see. 
                ////All false == Regular User
                //if (username.Text.Equals("CNS\\Sage.s.porter", StringComparison.InvariantCultureIgnoreCase))
                //{
                //    DomainAdmin = false;
                //    D365user = false;
                //    ReturnUser = true;
                //}

                //if (username.Text.Equals("CNS\\Rommel.McClaney", StringComparison.InvariantCultureIgnoreCase))
                //    launchChatToolStripMenuItem.Visible = true;
                if (DomainAdmin)
                {
                    if (username.Text.Equals("CNS\\Sage.s.porter", StringComparison.InvariantCultureIgnoreCase) || username.Text.Equals("CNS\\Mike.Sisson", StringComparison.InvariantCultureIgnoreCase))
                        launchChatToolStripMenuItem.Visible = true;
                    button1.Visible = true;
                    button2.Visible = true;
                    button5.Visible = true;
                    Delete.Visible = true;
                    Admin.Visible = true;
                    radioButton1.Visible = true;
                    radioButton2.Visible = true;
                    radioButton3.Visible = true;
                    radioButton4.Visible = true;
                    button9.Visible = true;
                    button10.Visible = true;
                    //save.Visible = true;
                    Returns.Visible = true;
                    label5.Text = "Domain Admins";

                    newReturnToolStripMenuItem.Visible = true;
                    newAdminToolStripMenuItem.Visible = true;
                    updateTicketToolStripMenuItem.Visible = true;
                    deleteTicketToolStripMenuItem.Visible = true;
                    saveTicketsToolStripMenuItem.Visible = true;
                    closeTicketToolStripMenuItem.Visible = true;
                    changeTabToolStripMenuItem.Visible = true;
                    saveByTicketToolStripMenuItem.Visible = true;

                }
                //D365
                if (D365user)
                {
                        D365user = true;
                        button5.Visible = true;
                        button1.Visible = true;
                        button2.Visible = true;
                        label5.Text = "D365 Support";

                    updateTicketToolStripMenuItem.Visible = true;
                    deleteTicketToolStripMenuItem.Visible = false;
                    closeTicketToolStripMenuItem.Visible = false;
                }
                //Returns
                if(ReturnUser&&!DomainAdmin)
                {
                    newReturnToolStripMenuItem.Visible = true;

                    radioButton1.Visible = true;
                        radioButton4.Visible = true;
                        button9.Visible = true;
                        button10.Visible = true;
                        Returns.Visible = true;
                        label5.Text = "Returns Support";
                }

                //givesaccess to guests for specific things. 
                AdditionalAccessDataContext vrs = new AdditionalAccessDataContext();
                var r = from p in vrs.AdditionalAccesses
                        select p;
                foreach (var x in r)
                {
                    if (x.GuestName == System.Security.Principal.WindowsIdentity.GetCurrent().Name)
                    {
                        if (x.GuestAccess.Contains("AdminAccess"))
                        {
                            button1.Visible = true;
                            button2.Visible = true;
                            button5.Visible = true;
                            Delete.Visible = true;
                            Admin.Visible = true;
                            radioButton1.Visible = true;
                            radioButton2.Visible = true;
                            radioButton3.Visible = true;
                            radioButton4.Visible = true;
                            button9.Visible = true;
                            button10.Visible = true;
                            //save.Visible = true;
                            Returns.Visible = true;
                            label5.Text = "Domain Admins";

                            newReturnToolStripMenuItem.Visible = true;
                            newAdminToolStripMenuItem.Visible = true;
                            updateTicketToolStripMenuItem.Visible = true;
                            deleteTicketToolStripMenuItem.Visible = true;
                            saveTicketsToolStripMenuItem.Visible = true;
                            closeTicketToolStripMenuItem.Visible = true;
                            changeTabToolStripMenuItem.Visible = true;
                            saveByTicketToolStripMenuItem.Visible = true;
                        }
                        if (x.GuestAccess.Contains("Chat"))
                        {
                            launchChatToolStripMenuItem.Visible = true;
                        }
                        if (x.GuestAccess.Contains("ReturnsAccess"))
                        {
                            newReturnToolStripMenuItem.Visible = true;

                            radioButton1.Visible = true;
                            radioButton4.Visible = true;
                            button9.Visible = true;
                            button10.Visible = true;
                            Returns.Visible = true;
                            label5.Text = "Returns Support";
                        }
                        if (x.GuestAccess.Contains("D365Access"))
                        {
                            D365user = true;
                            button5.Visible = true;
                            button1.Visible = true;
                            button2.Visible = true;
                            label5.Text = "D365 Support";

                            updateTicketToolStripMenuItem.Visible = true;
                            deleteTicketToolStripMenuItem.Visible = false;
                            closeTicketToolStripMenuItem.Visible = false;
                        }
                    }
                }

                LoadData();
            }
            catch (Exception err)
            {
                MessageBox.Show("We are sorry, the program cannot be ran at this time. Please check Internet and VPN connection.\r\n\r\n" + err.ToString());

                System.Diagnostics.Process cmd = new System.Diagnostics.Process();
                cmd.StartInfo.FileName = "cmd.exe";
                cmd.StartInfo.RedirectStandardInput = true;
                cmd.StartInfo.RedirectStandardOutput = true;
                cmd.StartInfo.CreateNoWindow = true;
                cmd.StartInfo.UseShellExecute = false;
                cmd.Start();

                cmd.StandardInput.WriteLine("ipconfig /flushdns");
                cmd.StandardInput.Flush();
                cmd.StandardInput.Close();
                cmd.WaitForExit();
                this.Close();
            }
            if (username.Text.Equals("CNS\\Mike.Sisson", StringComparison.InvariantCultureIgnoreCase) || username.Text.Equals("CNS\\Sage.s.porter", StringComparison.InvariantCultureIgnoreCase))
            {
                var chats = new chat();
                chats.Show();
            }
        }

        //Loads data into Listview
        private void addFOPAd(Ticketer x, bool changme)
        {
            string temp = "";
            string priority = "";
            if (x.Status == 'P')
            {
                if (changme == false)
                    counter += 1;
                scounter++;
                temp = "Pending Resolution";
                priority = checkpriority(x.Priority);
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
            if (x.Status == 'O')
            {
                if (changme == false)
                    counter += 1;
                scounter += 1;
                temp = "Open";
                priority = checkpriority(x.Priority);
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, x.Assigned };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
            if (x.Status == 'C')
            {

                DateTime dissue = new DateTime();
                dissue = Convert.ToDateTime(System.Convert.ToString(x.Dissue));
                DateTime dnow = new DateTime();
                dnow = DateTime.Now;

                TimeSpan duration = dnow - dissue;
                if (duration.TotalDays <= 7)
                {
                    scounter++;
                    if (changme == false)
                        counter += 1;
                    temp = "Closed";
                    priority = "";
                    string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                    listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                }
            }
            if (x.Status == 'F')
            {
                if (changme == false)
                    counter += 1;
                scounter += 1;
                temp = "Follow-Up Needed";
                priority = checkpriority(x.Priority);
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
        }
        private void addFOP(Ticketer x)
        {
            string temp = "";
            string priority = "";
            if (x.Status == 'P')
            {
                temp = "Pending Resolution";
                priority = checkpriority(x.Priority);
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
            if (x.Status == 'O')
            {
                temp = "Open";
                priority = checkpriority(x.Priority);
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, x.Assigned };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
            if (x.Status == 'C')
            {
                DateTime dissue = new DateTime();
                dissue = Convert.ToDateTime(System.Convert.ToString(x.Dissue));
                DateTime dnow = new DateTime();
                dnow = DateTime.Now;

                TimeSpan duration = dnow - dissue;
                if (duration.TotalDays <= 7)
                {
                    temp = "Closed";
                    priority = "";
                    string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                    listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                }
            }
            if (x.Status == 'F')
            {
                temp = "Follow-Up Needed";
                priority = checkpriority(x.Priority);
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
        }

        //overload methods for tickets that are closed on closed tab
        private void addFOPAd(Ticketer x, string closed)
        {
            string temp = "";
            string priority = "";
            if (x.Status == 'P')
            {
                if (ccount == false)
                    counter += 1;
                scounter++;
                temp = "Pending Resolution";
                priority = checkpriority(x.Priority);
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
            if (x.Status == 'O')
            {
                if (ccount == false)
                    counter += 1;
                scounter += 1;
                temp = "Open";
                priority = checkpriority(x.Priority);
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, x.Assigned };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
            if (x.Status == 'C')
            {

                DateTime dissue = new DateTime();
                dissue = Convert.ToDateTime(System.Convert.ToString(x.Dissue));
                DateTime dnow = new DateTime();
                dnow = DateTime.Now;

                TimeSpan duration = dnow - dissue;
                scounter++;
                if (ccount == false)
                    counter += 1;
                temp = "Closed";
                priority = "";
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
            if (x.Status == 'F')
            {
                if (ccount == false)
                    counter += 1;
                scounter += 1;
                temp = "Follow-Up Needed";
                priority = checkpriority(x.Priority);
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
        }
        private void addFOP(Ticketer x, string closed)
        {
            string temp = "";
            string priority = "";
            if (x.Status == 'P')
            {
                temp = "Pending Resolution";
                priority = checkpriority(x.Priority);
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
            if (x.Status == 'O')
            {
                temp = "Open";
                priority = checkpriority(x.Priority);
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, x.Assigned };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
            if (x.Status == 'C')
            {
                DateTime dissue = new DateTime();
                dissue = Convert.ToDateTime(System.Convert.ToString(x.Dissue));
                DateTime dnow = new DateTime();
                dnow = DateTime.Now;

                TimeSpan duration = dnow - dissue;
                temp = "Closed";
                priority = "";
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
            if (x.Status == 'F')
            {
                temp = "Follow-Up Needed";
                priority = checkpriority(x.Priority);
                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
            }
        }
        private void checkversion()
        {
            label7.Text = "Version "+version.ToString();

            ITVersionDataContext vrs = new ITVersionDataContext();
            var r = from p in vrs.ITTVersions
                    select p;
            foreach (var x in r)
            {
                if (version < x.Id)
                {
                    string cmdtext = "gpupdate /force";
                    System.Diagnostics.Process cmd = new System.Diagnostics.Process();
                    cmd.StartInfo.FileName = "cmd.exe";
                    cmd.StartInfo.RedirectStandardInput = true;
                    cmd.StartInfo.RedirectStandardOutput = true;
                    cmd.StartInfo.CreateNoWindow = true;
                    cmd.StartInfo.UseShellExecute = false;
                    cmd.Start();

                    cmd.StandardInput.WriteLine(cmdtext);
                    cmd.StandardInput.Flush();
                    cmd.StandardInput.Close();
                    cmd.WaitForExit();
                    cmd.Close();
                    MessageBox.Show(x.Messagetoshow+"\r\n\r\nYour Version is: "+version+"\r\nMost recent version: "+x.Id);
                    this.Close();
                }
            }
        }
        private void sorter(char sorts)
        {
            if (sorts == 'A')
            {
                listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
                listView1.Sort();
                this.listView1.ListViewItemSorter = new ListViewItemComparer(sortColumn, listView1.Sorting);
            }
            else if (sorts == 'D')
            {
                listView1.Sorting = System.Windows.Forms.SortOrder.Descending;
                listView1.Sort();
                this.listView1.ListViewItemSorter = new ListViewItemComparer(sortColumn, listView1.Sorting);
            }
            else
            {
                listView1.Sorting = System.Windows.Forms.SortOrder.Descending;
                listView1.Sort();
                sortColumn = 5;
                this.listView1.ListViewItemSorter = new ListViewItemComparer(5, listView1.Sorting);
            }
        }

        //pull data by radiobox selection.
        public void LoadData()
        {
            radioButton1.Checked = true;
            try
            {
                //Counting tickets if it is selected, change it to true and turn all others false. Clean counter
                if (dcount == false)
                {
                    scounter = 0;
                    counter = 0;
                    adcount = false;
                    ncount = false;
                    ccount = false;
                }
                else if (dcount==true)
                {
                    scounter = 0;
                }

                char sorts = ' ';
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Ascending)
                {
                    listView1.Sorting = System.Windows.Forms.SortOrder.None;
                    sorts = 'A';
                }
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Descending)
                {
                    sorts = 'D';
                    listView1.Sorting = System.Windows.Forms.SortOrder.None;
                }

                listView1.Sort();
                listView1.Items.Clear();

                checkversion();

                TicketsDBDataContext db = new TicketsDBDataContext();

                if (DomainAdmin == true)
                {
                    var r = from p in db.Ticketers
                            select p;

                    foreach (var x in r)
                    {
                        addFOPAd(x, dcount);
                    }
                    dcount = true;

                    //equipment returns
                    //-----------------------------------------------------------------------------------------
                    string temp = "";
                    string temp2 = "";
                    EquipReDataContext edb = new EquipReDataContext();
                    var r2 = from p in edb.EquipmentReturns
                             select p;
                    foreach (var x in r2)
                    {
                        if (x.TicketStatus == 'P')
                            temp = "Pending Resolution";
                        else if (x.TicketStatus == 'O')
                            temp = "Open";
                        else if (x.TicketStatus == 'C')
                            temp = "Closed";
                        else if (x.TicketStatus == 'F')
                            temp = "Follow-Up";

                        if (x.Priority == 'L')
                            temp2 = "Low";
                        else if (x.Priority == 'M')
                            temp2 = "Medium";
                        else if (x.Priority == 'H')
                            temp2 = "High";

                        DateTime dnow = new DateTime();
                        dnow = DateTime.Now;
                        int duration = (int)(dnow - x.DateCreated).Value.TotalDays;
                        if (x.TicketStatus == 'C')
                        {
                            if (duration<=7)
                            {
                                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.DateCreated), "Return " + x.TagID, temp2, "" };
                                listView1.Items.Add(System.Convert.ToString(x.IDText)).SubItems.AddRange(row);
                            }
                        }
                        else
                        {
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.DateCreated), "Return " + x.TagID, temp2, "" };
                            listView1.Items.Add(System.Convert.ToString(x.IDText)).SubItems.AddRange(row);
                        }
                    }
                    //-----------------------------------------------------------------------------------------
                    sorter(sorts);
                }
                if (D365user == true)
                {
                    var r = from p in db.Ticketers.Where<Ticketer>(x => x.Ticketor.ToString() == username.Text || x.IssueType.ToString() == "D365")
                            select p;

                    foreach (var x in r)
                    {
                        addFOP(x);
                    }
                    dcount = true;
                    sorter(sorts);
                }
                if(ReturnUser==true && DomainAdmin!=true)
                {
                    var r = from p in db.Ticketers.Where<Ticketer>(x => x.Ticketor.ToString() == username.Text)
                            select p;
                    foreach (var x in r)
                    {
                        addFOP(x);
                    }
                    //equipment returns
                    //-----------------------------------------------------------------------------------------
                    string temp = "";
                    string temp2 = "";
                    EquipReDataContext edb = new EquipReDataContext();
                    var r2 = from p in edb.EquipmentReturns
                             select p;
                    foreach (var x in r2)
                    {
                        if (x.TicketStatus == 'P')
                            temp = "Pending Resolution";
                        else if (x.TicketStatus == 'O')
                            temp = "Open";
                        else if (x.TicketStatus == 'C')
                            temp = "Closed";
                        else if (x.TicketStatus == 'F')
                            temp = "Follow-Up";

                        if (x.Priority == 'L')
                            temp2 = "Low";
                        else if (x.Priority == 'M')
                            temp2 = "Medium";
                        else if (x.Priority == 'H')
                            temp2 = "High";

                        DateTime dnow = new DateTime();
                        dnow = DateTime.Now;
                        int duration = (int)(dnow - x.DateCreated).Value.TotalDays;
                        if (x.TicketStatus == 'C')
                        {
                            if (duration <= 7)
                            {
                                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.DateCreated), "Return " + x.TagID, temp2, "" };
                                listView1.Items.Add(System.Convert.ToString(x.IDText)).SubItems.AddRange(row);
                            }
                        }
                        else
                        {
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.DateCreated), "Return " + x.TagID, temp2, "" };
                            listView1.Items.Add(System.Convert.ToString(x.IDText)).SubItems.AddRange(row);
                        }
                    }
                    //-----------------------------------------------------------------------------------------
                    sorter(sorts);
                }
                if (DomainAdmin == false && D365user == false&& ReturnUser==false)
                {


                    var r = from p in db.Ticketers
                            where p.Ticketor.ToString() == username.Text
                            select p;

                    foreach (var x in r)
                    {
                        addFOP(x);
                    }
                    sorter(sorts);
                }
                if (counter < scounter)
                {
                    Displaynotify();
                    FlashWindow.Flash(this);
                    counter = scounter;
                }
            }
            catch
            {
                MessageBox.Show("We are sorry, we are not able to run the program at this time. Please check Internet and VPN connections.");
            }
        }
        public void LoadAdminData()
        {
            try
            {
                //Counting tickets if it is selected, change it to true and turn all others false. Clean counter
                if (adcount == false)
                {
                    scounter = 0;
                    counter = 0;
                    dcount = false;
                    ncount = false;
                    ccount = false;
                }
                else if (adcount == true)
                {
                    scounter = 0;
                }
                char sorts = ' ';
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Ascending)
                {
                    listView1.Sorting = System.Windows.Forms.SortOrder.None;
                    sorts = 'A';
                }
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Descending)
                {
                    sorts = 'D';
                    listView1.Sorting = System.Windows.Forms.SortOrder.None;
                }

                listView1.Sort();
                listView1.Items.Clear();

                checkversion();

                TicketsDBDataContext db = new TicketsDBDataContext();

                if (DomainAdmin == true)
                {
                    var r = from p in db.Ticketers.Where<Ticketer>(x =>
                            System.Convert.ToString(x.IssueType).Contains("IT")
                            )
                            select p;

                    foreach (var x in r)
                    {
                        addFOPAd(x, adcount);
                    }
                    adcount = true;
                    sorter(sorts);
                }
                if (D365user == true)
                {


                    var r = from p in db.Ticketers.Where<Ticketer>(x => x.Ticketor.ToString() == username.Text || x.IssueType.ToString() == "D365")
                            select p;

                    foreach (var x in r)
                    {
                        addFOP(x);
                    }
                    adcount = true;
                    sorter(sorts);
                }
                if (DomainAdmin == false && D365user == false)
                {


                    var r = from p in db.Ticketers
                            where p.Ticketor.ToString() == username.Text
                            select p;

                    foreach (var x in r)
                    {
                        addFOP(x);
                    }
                    adcount = true;
                    sorter(sorts);
                }
                if (counter < scounter)
                {
                    Displaynotify();
                    FlashWindow.Flash(this);
                    counter = scounter;
                }
            }
            catch
            {
                MessageBox.Show("We are sorry, we are not able to run the program at this time. Please check Internet and VPN connections.");
            }
        }
        public void LoadNonAdmin()
        {
            try
            {
                //Counting tickets if it is selected, change it to true and turn all others false. Clean counter
                if (ncount == false)
                {
                    scounter = 0;
                    counter = 0;
                    adcount = false;
                    dcount = false;
                    ccount = false;
                }
                else if (ncount == true)
                {
                    scounter = 0;
                }
                char sorts = ' ';
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Ascending)
                {
                    listView1.Sorting = System.Windows.Forms.SortOrder.None;
                    sorts = 'A';
                }
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Descending)
                {
                    sorts = 'D';
                    listView1.Sorting = System.Windows.Forms.SortOrder.None;
                }

                listView1.Sort();
                listView1.Items.Clear();

                checkversion();

                TicketsDBDataContext db = new TicketsDBDataContext();

                if (DomainAdmin == true)
                {
                    var r = from p in db.Ticketers.Where<Ticketer>(x =>
                            (!System.Convert.ToString(x.IssueType).Contains("IT"))&&(!System.Convert.ToString(x.IssueType).Contains("D365"))
                            )
                            select p;

                    foreach (var x in r)
                    {
                        addFOPAd(x, ncount);
                    }
                    ncount = true;
                    sorter(sorts);
                }
                if (D365user == true)
                {
                    var r = from p in db.Ticketers.Where<Ticketer>(x => x.Ticketor.ToString() == username.Text || x.IssueType.ToString() == "D365")
                            select p;

                    foreach (var x in r)
                    {
                        addFOP(x);
                    }
                    ncount = true;
                    sorter(sorts);
                }
                if (DomainAdmin == false && D365user == false)
                {
                    var r = from p in db.Ticketers
                            where p.Ticketor.ToString() == username.Text
                            select p;

                    foreach (var x in r)
                    {
                        addFOP(x);
                    }
                    ncount = true;
                    sorter(sorts);
                }
                if (counter < scounter)
                {
                    Displaynotify();
                    FlashWindow.Flash(this);
                    counter = scounter;
                }
            }
            catch
            {
                MessageBox.Show("We are sorry, we are not able to run the program at this time. Please check Internet and VPN connections.");
            }
        }
        public void LoadAllClosed()
        {
            try
            {
                //Counting tickets if it is selected, change it to true and turn all others false. Clean counter
                if (ccount == false)
                {
                    scounter = 0;
                    counter = 0;
                    adcount = false;
                    ncount = false;
                    dcount = false;
                }
                else if (ccount == true)
                {
                    scounter = 0;
                }
                char sorts = ' ';
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Ascending)
                {
                    listView1.Sorting = System.Windows.Forms.SortOrder.None;
                    sorts = 'A';
                }
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Descending)
                {
                    sorts = 'D';
                    listView1.Sorting = System.Windows.Forms.SortOrder.None;
                }

                listView1.Sort();
                listView1.Items.Clear();

                checkversion();

                TicketsDBDataContext db = new TicketsDBDataContext();

                if (DomainAdmin == true)
                {


                    var r = from p in db.Ticketers.Where<Ticketer>(x =>
                            x.Status=='C'
                            )
                            select p;

                    foreach (var x in r)
                    {
                        addFOPAd(x, "Closed");
                    }
                    ccount = true;

                    //Closed returns
//-----------------------------------------------------------------------------------------
                    string temp = "";
                    string temp2 = "";
                    EquipReDataContext edb = new EquipReDataContext();
                    var r2 = from p in edb.EquipmentReturns
                             select p;
                    foreach (var x in r2)
                    {
                        if (x.TicketStatus == 'P')
                            temp = "Pending Resolution";
                        else if (x.TicketStatus == 'O')
                            temp = "Open";
                        else if (x.TicketStatus == 'C')
                            temp = "Closed";
                        else if (x.TicketStatus == 'F')
                            temp = "Follow-Up";

                        if (x.Priority == 'L')
                            temp2 = "Low";
                        else if (x.Priority == 'M')
                            temp2 = "Medium";
                        else if (x.Priority == 'H')
                            temp2 = "High";

                        if (x.TicketStatus == 'C')
                        {
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.DateCreated), "Return " + x.TagID, temp2, "" };
                            listView1.Items.Add(System.Convert.ToString(x.IDText)).SubItems.AddRange(row);
                        }
                    }
//-----------------------------------------------------------------------------------------
                    sorter(sorts);
                }
                if (D365user == true)
                {
                    var r = from p in db.Ticketers.Where<Ticketer>(x => x.Ticketor.ToString() == username.Text || x.IssueType.ToString() == "D365")
                            select p;

                    foreach (var x in r)
                    {
                        addFOP(x, "Closed");
                    }
                    ccount = true;
                    sorter(sorts);
                }
                if (ReturnUser=true && DomainAdmin==false)
                {
                    var r = from p in db.Ticketers.Where<Ticketer>(x => x.Ticketor.ToString() == username.Text)
                            select p;

                    foreach (var x in r)
                    {
                        addFOP(x, "Closed");
                    }
                    //Closed returns
                    //-----------------------------------------------------------------------------------------
                    string temp = "";
                    string temp2 = "";
                    EquipReDataContext edb = new EquipReDataContext();
                    var r2 = from p in edb.EquipmentReturns
                             select p;
                    foreach (var x in r2)
                    {
                        if (x.TicketStatus == 'P')
                            temp = "Pending Resolution";
                        else if (x.TicketStatus == 'O')
                            temp = "Open";
                        else if (x.TicketStatus == 'C')
                            temp = "Closed";
                        else if (x.TicketStatus == 'F')
                            temp = "Follow-Up";

                        if (x.Priority == 'L')
                            temp2 = "Low";
                        else if (x.Priority == 'M')
                            temp2 = "Medium";
                        else if (x.Priority == 'H')
                            temp2 = "High";

                        if (x.TicketStatus == 'C')
                        {
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.DateCreated), "Return " + x.TagID, temp2, "" };
                            listView1.Items.Add(System.Convert.ToString(x.IDText)).SubItems.AddRange(row);
                        }
                    }
                    //-----------------------------------------------------------------------------------------
                }
                if (DomainAdmin == false && D365user == false && ReturnUser==false)
                {


                    var r = from p in db.Ticketers
                            where p.Ticketor.ToString() == username.Text
                            select p;

                    foreach (var x in r)
                    {
                        addFOP(x, "Closed");
                    }
                    ccount = true;
                    sorter(sorts);
                }
                if (counter < scounter)
                {
                    Displaynotify();
                    FlashWindow.Flash(this);
                    counter = scounter;
                }
            }
            catch
            {
                MessageBox.Show("We are sorry, we are not able to run the program at this time. Please check Internet and VPN connections.");
            }
        }

        //Checks priority of tickets and displays it.
        private string checkpriority(char? CPriority)
        {
            string sPriority = "";
            if (CPriority == 'L')
                sPriority = "Low";
            if (CPriority == 'M')
                sPriority = "Medium";
            if (CPriority == 'H')
                sPriority = "High";
            if (CPriority == 'P')
                sPriority = "Low";
            return sPriority;
        }

        //Server Data refresh. - 15 minutes
        private void timer_tick(object sender, EventArgs e)
        {
            username.Text = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            label3.Text = DateTime.Now.ToString("MM-dd-yyyy h:mm tt");
            if(radioButton1.Checked)
                LoadData();
            if (radioButton2.Checked)
                LoadAdminData();
            if (radioButton3.Checked)
                LoadNonAdmin();
            if (Returns.Checked)
                loadreturns();
        }


        //only for equipment tickets
        public void loadreturns()
        {
            try
            {
                char sorts = ' ';
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Ascending)
                {
                    listView1.Sorting = System.Windows.Forms.SortOrder.None;
                    sorts = 'A';
                }
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Descending)
                {
                    sorts = 'D';
                    listView1.Sorting = System.Windows.Forms.SortOrder.None;
                }
                listView1.Sort();
                listView1.Items.Clear();
                string temp = "";
                string temp2 = "";
                EquipReDataContext edb = new EquipReDataContext();
                var r2 = from p in edb.EquipmentReturns
                         select p;
                foreach (var x in r2)
                {
                    if (x.TicketStatus == 'P')
                        temp = "Pending Resolution";
                    else if (x.TicketStatus == 'O')
                        temp = "Open";
                    else if (x.TicketStatus == 'C')
                        temp = "Closed";
                    else if (x.TicketStatus == 'F')
                        temp = "Follow-Up";

                    if (x.Priority == 'L')
                        temp2 = "Low";
                    else if (x.Priority == 'M')
                        temp2 = "Medium";
                    else if (x.Priority == 'H')
                        temp2 = "High";

                    DateTime dnow = new DateTime();
                    dnow = DateTime.Now;
                    int duration = (int)(dnow - x.DateCreated).Value.TotalDays;
                    if (x.TicketStatus == 'C')
                    {
                        if (duration <= 7)
                        {
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.DateCreated),"Return "+ x.TagID, temp2, "" };
                            listView1.Items.Add(System.Convert.ToString(x.IDText)).SubItems.AddRange(row);
                        }
                    }
                    else
                    {
                        string[] row = { temp, x.Ticketor, System.Convert.ToString(x.DateCreated), "Return " + x.TagID, temp2, "" };
                        listView1.Items.Add(System.Convert.ToString(x.IDText)).SubItems.AddRange(row);
                    }
                }
                sorter(sorts);
            }
            catch
            {
                MessageBox.Show("We are sorry, we are not able to run the program at this time. Please check Internet and VPN connections.");
            }
        }
        #endregion

        #region Buttons
        private void button10_Click(object sender, EventArgs e)
        {
            if (listView1.FocusedItem != null)
            {
                if (listView1.FocusedItem.Text.Contains("E"))
                {
                    var returndatas = new ReturnData(listView1.FocusedItem.Text, this, username.Text);
                    returndatas.Show();
                }
                else
                    MessageBox.Show("Please select a ticket that is equipment based");
            }
            else
                MessageBox.Show("Please select a ticket that is equipment based");
        }
        //New Return button
        private void button9_Click(object sender, EventArgs e)
        {
            var returns = new NewReturn(username.Text, this);
            returns.Show();
        }

        ////chat button
        //private void chat_Click(object sender, EventArgs e)
        //{
        //    var chats = new chat();
        //    chats.Show();
        //}

        //admin Button
        private void Admin_Click(object sender, EventArgs e)
        {
            var admin = new AdminTickets(this);
            admin.Show();
        }

        //Save Button. 
        private void save_Click(object sender, EventArgs e)
        {
            var Savedata = new Savedates();
            Savedata.Show();
        }

        //help button
        private void button8_Click(object sender, EventArgs e)
        {
            var MD = new Form3();
            MD.Show();
        }

        //Search Button
        private void Search(object sender, EventArgs e)
        {
            if (textBox1.Text != "Username, Status, Date, ID")
            {
                textBox1.Text = textBox1.Text.Replace(" ", "");
                char sorts = ' ';
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Ascending)
                {
                    listView1.Sorting = System.Windows.Forms.SortOrder.None;
                    sorts = 'A';
                }
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Descending)
                {
                    sorts = 'D';
                    listView1.Sorting = System.Windows.Forms.SortOrder.None;
                }
                listView1.Sort();
                listView1.Items.Clear();
                TicketsDBDataContext db = new TicketsDBDataContext();
                EquipReDataContext ret = new EquipReDataContext();
                if (DomainAdmin == true)
                {
                    var r = from p in db.Ticketers.Where<Ticketer>(x =>
                            System.Convert.ToString(x.Ticketor).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.ID2).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.Dissue).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.Status).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.IssueType).Contains(textBox1.Text)
                            )
                            select p;
                    string priority = "";
                    foreach (var x in r)
                    {

                        string temp = "";
                        if (x.Status == 'P')
                        {
                            temp = "Pending Resolution";
                            priority = checkpriority(x.Priority);
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                            listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                        if (x.Status == 'O')
                        {
                            temp = "Open";
                            priority = checkpriority(x.Priority);
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, x.Assigned };
                            listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                        if (x.Status == 'C')
                        {
                                temp = "Closed";
                                priority = "";
                                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                        if (x.Status == 'F')
                        {
                            temp = "Follow-Up Needed";
                            priority = checkpriority(x.Priority);
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                            listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                    }
                    if (sorts == 'A')
                    {
                        listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
                        listView1.Sort();
                        this.listView1.ListViewItemSorter = new ListViewItemComparer(sortColumn, listView1.Sorting);
                    }
                    else if (sorts == 'D')
                    {
                        listView1.Sorting = System.Windows.Forms.SortOrder.Descending;
                        listView1.Sort();
                        this.listView1.ListViewItemSorter = new ListViewItemComparer(sortColumn, listView1.Sorting);
                    }
                    else
                    {
                        listView1.Sorting = System.Windows.Forms.SortOrder.Descending;
                        listView1.Sort();
                        sortColumn = 5;
                        this.listView1.ListViewItemSorter = new ListViewItemComparer(5, listView1.Sorting);
                    }
                }
                else if (DomainAdmin == false && D365user == true && ReturnUser == false)
                {
                    var r = from p in db.Ticketers.Where<Ticketer>(x => ((
                            System.Convert.ToString(x.Ticketor).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.ID2).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.Dissue).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.Status).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.IssueType).Contains(textBox1.Text)) && x.Ticketor.ToString().Contains(username.Text))|| ((
                            System.Convert.ToString(x.Ticketor).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.ID2).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.Dissue).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.Status).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.IssueType).Contains(textBox1.Text)) && x.IssueType.Contains("D365"))
                            )
                            select p;
                    string priority = "";
                    foreach (var x in r)
                    {

                        string temp = "";
                        if (x.Status == 'P')
                        {
                            temp = "Pending Resolution";
                            priority = checkpriority(x.Priority);
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                            listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                        if (x.Status == 'O')
                        {
                            temp = "Open";
                            priority = checkpriority(x.Priority);
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, x.Assigned };
                            listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                        if (x.Status == 'C')
                        {
                            DateTime dissue = new DateTime();
                            dissue = Convert.ToDateTime(System.Convert.ToString(x.Dissue));
                            DateTime dnow = new DateTime();
                            dnow = DateTime.Now;

                            TimeSpan duration = dnow - dissue;
                            if (duration.TotalDays <= 7)
                            {
                                temp = "Closed";
                                priority = "";
                                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                            }
                        }
                        if (x.Status == 'F')
                        {
                            temp = "Follow-Up Needed";
                            priority = checkpriority(x.Priority);
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                            listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                    }
                    if (sorts == 'A')
                    {
                        listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
                        listView1.Sort();
                        this.listView1.ListViewItemSorter = new ListViewItemComparer(sortColumn, listView1.Sorting);
                    }
                    else if (sorts == 'D')
                    {
                        listView1.Sorting = System.Windows.Forms.SortOrder.Descending;
                        listView1.Sort();
                        this.listView1.ListViewItemSorter = new ListViewItemComparer(sortColumn, listView1.Sorting);
                    }
                    else
                    {
                        listView1.Sorting = System.Windows.Forms.SortOrder.Descending;
                        listView1.Sort();
                        sortColumn = 5;
                        this.listView1.ListViewItemSorter = new ListViewItemComparer(5, listView1.Sorting);
                    }
                }
                else if (DomainAdmin == false && D365user == false && ReturnUser == true)
                {
                    var r = from p in db.Ticketers.Where<Ticketer>(x => ((
                            System.Convert.ToString(x.Ticketor).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.ID2).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.Dissue).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.Status).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.IssueType).Contains(textBox1.Text)) && x.Ticketor.ToString().Contains(username.Text))
                            )
                            select p;
                    EquipReDataContext edb = new EquipReDataContext();
                    var eq = from p in edb.EquipmentReturns.Where<EquipmentReturn>(x =>(x.IDText.Contains(textBox1.Text)||
                             x.TagID.Contains(textBox1.Text) || x.DateCreated.ToString().Contains(textBox1.Text)||
                             x.Ticketor.Contains(textBox1.Text)||x.TicketStatus.ToString().Contains(textBox1.Text)))
                            select p;

                    string priority = "";
                    foreach (var x in r)
                    {

                        string temp = "";
                        if (x.Status == 'P')
                        {
                            temp = "Pending Resolution";
                            priority = checkpriority(x.Priority);
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                            listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                        if (x.Status == 'O')
                        {
                            temp = "Open";
                            priority = checkpriority(x.Priority);
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, x.Assigned };
                            listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                        if (x.Status == 'C')
                        {
                            DateTime dissue = new DateTime();
                            dissue = Convert.ToDateTime(System.Convert.ToString(x.Dissue));
                            DateTime dnow = new DateTime();
                            dnow = DateTime.Now;

                            TimeSpan duration = dnow - dissue;
                            if (duration.TotalDays <= 7)
                            {
                                temp = "Closed";
                                priority = "";
                                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                            }
                        }
                        if (x.Status == 'F')
                        {
                            temp = "Follow-Up Needed";
                            priority = checkpriority(x.Priority);
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                            listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                    }
                    //returns search
                    string temp1 = "";
                    string temp2 = "";
                    foreach (var x in eq)
                    {
                        if (x.TicketStatus == 'P')
                            temp1 = "Pending Resolution";
                        else if (x.TicketStatus == 'O')
                            temp1 = "Open";
                        else if (x.TicketStatus == 'C')
                            temp1 = "Closed";
                        else if (x.TicketStatus == 'F')
                            temp1 = "Follow-Up";

                        if (x.Priority == 'L')
                            temp2 = "Low";
                        else if (x.Priority == 'M')
                            temp2 = "Medium";
                        else if (x.Priority == 'H')
                            temp2 = "High";

                        DateTime dnow = new DateTime();
                        dnow = DateTime.Now;
                        int duration = (int)(dnow - x.DateCreated).Value.TotalDays;
                        if (x.TicketStatus == 'C')
                        {
                            if (duration <= 7)
                            {
                                string[] row = { temp1, x.Ticketor, System.Convert.ToString(x.DateCreated), "Return " + x.TagID, temp2, "" };
                                listView1.Items.Add(System.Convert.ToString(x.IDText)).SubItems.AddRange(row);
                            }
                        }
                        else
                        {
                            string[] row = { temp1, x.Ticketor, System.Convert.ToString(x.DateCreated), "Return " + x.TagID, temp2, "" };
                            listView1.Items.Add(System.Convert.ToString(x.IDText)).SubItems.AddRange(row);
                        }
                    }
                    if (sorts == 'A')
                    {
                        listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
                        listView1.Sort();
                        this.listView1.ListViewItemSorter = new ListViewItemComparer(sortColumn, listView1.Sorting);
                    }
                    else if (sorts == 'D')
                    {
                        listView1.Sorting = System.Windows.Forms.SortOrder.Descending;
                        listView1.Sort();
                        this.listView1.ListViewItemSorter = new ListViewItemComparer(sortColumn, listView1.Sorting);
                    }
                    else
                    {
                        listView1.Sorting = System.Windows.Forms.SortOrder.Descending;
                        listView1.Sort();
                        sortColumn = 5;
                        this.listView1.ListViewItemSorter = new ListViewItemComparer(5, listView1.Sorting);
                    }
                }

                else if (DomainAdmin == false && D365user == false && ReturnUser==false)
                {
                    var r = from p in db.Ticketers.Where<Ticketer>(x => (
                            System.Convert.ToString(x.Ticketor).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.ID2).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.Dissue).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.Status).Contains(textBox1.Text) ||
                            System.Convert.ToString(x.IssueType).Contains(textBox1.Text)) && x.Ticketor.ToString() == username.Text
                            )
                            select p;
                    string priority = "";
                    foreach (var x in r)
                    {

                        string temp = "";
                        if (x.Status == 'P')
                        {
                            temp = "Pending Resolution";
                            priority = checkpriority(x.Priority);
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                            listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                        if (x.Status == 'O')
                        {
                            temp = "Open";
                            priority = checkpriority(x.Priority);
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, x.Assigned };
                            listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                        if (x.Status == 'C')
                        {
                            DateTime dissue = new DateTime();
                            dissue = Convert.ToDateTime(System.Convert.ToString(x.Dissue));
                            DateTime dnow = new DateTime();
                            dnow = DateTime.Now;

                            TimeSpan duration = dnow - dissue;
                            if (duration.TotalDays <= 7)
                            {
                                temp = "Closed";
                                priority = "";
                                string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                                listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                            }
                        }
                        if (x.Status == 'F')
                        {
                            temp = "Follow-Up Needed";
                            priority = checkpriority(x.Priority);
                            string[] row = { temp, x.Ticketor, System.Convert.ToString(x.Dissue), x.IssueType, priority, "" };
                            listView1.Items.Add(System.Convert.ToString(x.ID2)).SubItems.AddRange(row);
                        }
                    }
                    if (sorts == 'A')
                    {
                        listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
                        listView1.Sort();
                        this.listView1.ListViewItemSorter = new ListViewItemComparer(sortColumn, listView1.Sorting);
                    }
                    else if (sorts == 'D')
                    {
                        listView1.Sorting = System.Windows.Forms.SortOrder.Descending;
                        listView1.Sort();
                        this.listView1.ListViewItemSorter = new ListViewItemComparer(sortColumn, listView1.Sorting);
                    }
                    else
                    {
                        listView1.Sorting = System.Windows.Forms.SortOrder.Descending;
                        listView1.Sort();
                        sortColumn = 5;
                        this.listView1.ListViewItemSorter = new ListViewItemComparer(5, listView1.Sorting);
                    }
                }
            }
        }

        //Refresh Button
        private void RefreshForm(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
                LoadData();
            if (radioButton2.Checked)
                LoadAdminData();
            if (radioButton3.Checked)
                LoadNonAdmin();
            if (Returns.Checked)
                loadreturns();
        }

        //ClearSearch Button
        private void ClearSearch(object sender, EventArgs e)
        {
            LoadData();
            textBox1.Text = "Username, Status, Date, ID";
        }

        //New Ticket Button
        private void NewTicket(object sender, EventArgs e)
        {
            var Moredet = new Form2(this, DomainAdmin);
            Moredet.Show();
        }

        //CloseTicket Button
        private void CloseTicket(object sender, EventArgs e)
        {
            if (!listView1.FocusedItem.Text.Contains("E"))
            {
                if (listView1.FocusedItem != null)
                {
                    TicketsDBDataContext db = new TicketsDBDataContext();

                    var r = from p in db.Ticketers
                            where p.ID2 == long.Parse(listView1.FocusedItem.Text)
                            select p;

                    foreach (var x in r)
                    {
                        if (x.Status != 'C')
                        {
                            long selectedID = long.Parse(listView1.FocusedItem.Text);
                            var update = new CloseTicket(selectedID, this);
                            update.Show();
                        }
                        else
                        {
                            MessageBox.Show("Ticket is already closed.");
                        }
                    }
                }
            }
            else if (listView1.FocusedItem.Text.Contains("E"))
            {
                    if (listView1.FocusedItem != null)
                    {
                        var returndatas = new ReturnData(listView1.FocusedItem.Text, this, username.Text);
                        returndatas.Show();
                    }
                    else
                    {
                        MessageBox.Show("Please select an Item to update!", "Invalid Selection",
        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
            }
            else
            {
                MessageBox.Show("Please select a ticket to close!", "Invalid Selection",
MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Update Button
        private void LaunchUpdate(object sender, EventArgs e)
        {
            if (!listView1.FocusedItem.Text.Contains("E"))
            {
                if (listView1.FocusedItem != null)
                {
                    long selectedID = long.Parse(listView1.FocusedItem.Text);
                    var update = new UpdateTicket(selectedID, this, DomainAdmin);
                    update.Show();
                }
                else
                {
                    MessageBox.Show("Please select an Item to update!", "Invalid Selection",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                if (listView1.FocusedItem != null)
                {
                    var returndatas = new ReturnData(listView1.FocusedItem.Text, this, username.Text);
                    returndatas.Show();
                }
                else
                {
                    MessageBox.Show("Please select an Item to update!", "Invalid Selection",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //MoreDetails Button
        private void LaunchMoreDetails(object sender, EventArgs e)
        {
            if (!listView1.FocusedItem.Text.Contains("E"))
            {
                if (listView1.FocusedItem != null)
                {
                    long selectedID = long.Parse(listView1.FocusedItem.Text);
                    var MD = new More_Details(selectedID);
                    MD.Show();
                }
                else
                {
                    MessageBox.Show("Please Select a item to review details", "Invalid Selection",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                if (listView1.FocusedItem != null)
                {
                    var returndatas = new ReturnData(listView1.FocusedItem.Text, this, username.Text);
                    returndatas.Show();
                }
                else
                {
                    MessageBox.Show("Please select an Item to update!", "Invalid Selection",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //Delete Button
        private void Delete_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to delete this entry?", "Delete", MessageBoxButtons.YesNo);
            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                counter -= 1;
                TicketsDBDataContext db = new TicketsDBDataContext();
                if (listView1.FocusedItem != null)
                {
                    var deleter =
                        from details in db.Ticketers
                        where details.ID2 == long.Parse(listView1.FocusedItem.Text)
                        select details;

                    foreach (var detail in deleter)
                    {
                        db.Ticketers.DeleteOnSubmit(detail);
                    }

                    try
                    {
                        db.SubmitChanges();
                    }
                    catch
                    {
                        Console.WriteLine("Unable to delete item");
                        // Provide for exceptions.
                    }
                    LoadData();
                }
                else
                {
                    MessageBox.Show("Please Select an Item to Delete", "Unable to delete Ticket",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        #endregion

        #region Shortcuts
        //CTRL+ shortcut commands
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            //shortcut key for new ticket (CTRL+N)
            if (keyData == (Keys.Control | Keys.N))
            {
                var Moredet = new Form2(this, DomainAdmin);
                Moredet.Show();
                return true;
            }

            //shortcut key for Update ticket (CTRL+U)
            else if (keyData == (Keys.Control | Keys.U))
            {
                if (DomainAdmin || D365user)
                {
                    if (listView1.FocusedItem != null)
                    {
                        long selectedID = long.Parse(listView1.FocusedItem.Text);
                        var update = new UpdateTicket(selectedID, this, DomainAdmin);
                        update.Show();
                    }
                    else
                    {
                        MessageBox.Show("Please select an Item to update!", "Invalid Selection",
        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return true;
                }
            }

            //shortcut key for Save data
            else if (keyData == (Keys.Control | Keys.S))
            {
                if (DomainAdmin)
                {
                    var Savedata = new Savedates();
                    Savedata.Show();
                }
            }

            //shortcut key for More details. 
            else if (keyData == (Keys.Control | Keys.D))
            {
                if (DomainAdmin || D365user)
                {
                    if (listView1.FocusedItem != null)
                    {
                        long selectedID = long.Parse(listView1.FocusedItem.Text);
                        var MD = new More_Details(selectedID);
                        MD.Show();
                    }
                    else
                    {
                        MessageBox.Show("Please Select a item to review details", "Invalid Selection",
        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            //shortcut key for Delete
            else if (keyData == (Keys.Delete))
            {
                if (DomainAdmin)
                {
                    EventArgs e = new EventArgs();
                    Delete_Click(this, e);
                }
            }

            //help key
            else if (keyData == (Keys.Control | Keys.H))
            {
                EventArgs e = new EventArgs();
                button8_Click(this, e);
            }

            //refresh key
            else if (keyData == (Keys.F5))
            {
                EventArgs e = new EventArgs();
                RefreshForm(this, e);
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        //Search While Typing Command
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Search(sender, e);
        }

        //Launch Details with double click
        private void Itemdouble(object sender, EventArgs e)
        {
            if (DomainAdmin == true)
            {
                foreach (ListViewItem LItem in listView1.SelectedItems)
                {
                    if (LItem.ToString().Contains("E"))
                    {
                        var returndatas = new ReturnData(listView1.FocusedItem.Text, this, username.Text);
                        returndatas.Show();
                    }
                    else
                    {
                        LaunchMoreDetails(sender, e);
                    }
                }
            }
        }

        //right click copy command
        private void Copy_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count != 0)
            {
                foreach (ListViewItem LItem in listView1.SelectedItems)
                {
                    Clipboard.SetText(LItem.Text);
                }
            }
        }

        //Clear searchbox on select
        private void searchclicked(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        //sort by column
        int sortColumn = -1;
        private void listview1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Determine whether the column is the same as the last column clicked.
            if (e.Column != sortColumn)
            {
                // Set the sort column to the new column.
                sortColumn = e.Column;
                // Set the sort order to ascending by default.
                listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
            }
            else
            {
                // Determine what the last sort order was and change it.
                if (listView1.Sorting == System.Windows.Forms.SortOrder.Ascending)
                    listView1.Sorting = System.Windows.Forms.SortOrder.Descending;
                else
                    listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
            }

            // Call the sort method to manually sort.
            listView1.Sort();
            // Set the ListViewItemSorter property to a new ListViewItemComparer
            // object.
            this.listView1.ListViewItemSorter = new ListViewItemComparer(e.Column, listView1.Sorting);
        }
        class ListViewItemComparer : IComparer
        {
            private int col;
            private System.Windows.Forms.SortOrder order;
            public ListViewItemComparer()
            {
                col = 0;
                order = System.Windows.Forms.SortOrder.Ascending;
            }
            public ListViewItemComparer(int column, System.Windows.Forms.SortOrder order)
            {
                col = column;
                this.order = order;
            }
            public int Compare(object x, object y)
            {
                int returnVal = -1;
                returnVal = String.Compare(((ListViewItem)x).SubItems[col].Text,
                                        ((ListViewItem)y).SubItems[col].Text);
                // Determine whether the sort order is descending.
                if (order == System.Windows.Forms.SortOrder.Descending)
                    // Invert the value returned by String.Compare.
                    returnVal *= -1;
                return returnVal;
            }
        }
        #endregion

        #region tabs
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                LoadData();
            }
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
                LoadAdminData();
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
                LoadNonAdmin();
        }
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
                LoadAllClosed();
        }
        private void Returns_CheckedChanged(object sender, EventArgs e)
        {
            if(Returns.Checked)
                loadreturns();
        }

        #endregion

        #region Menu
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void pictureBox1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.cnscares.com/");
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
                LoadData();
            if (radioButton2.Checked)
                LoadAdminData();
            if (radioButton3.Checked)
                LoadNonAdmin();
            if (Returns.Checked)
                loadreturns();
        }

        private void launchChatToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var chats = new chat();
            chats.Show();
        }

        private void saveTicketsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var Savedata = new Savedates();
            Savedata.Show();
        }

        private void moreHelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var MD = new Form3();
            MD.Show();
        }

        private void changeFontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (fontDialog1.ShowDialog() != DialogResult.Cancel)
            {
                label1.Font = fontDialog1.Font;
                label2.Font = fontDialog1.Font;
                label3.Font = fontDialog1.Font;
                label4.Font = fontDialog1.Font;
                label5.Font = fontDialog1.Font;
                username.Font = fontDialog1.Font;
                radioButton1.Font = fontDialog1.Font;
                radioButton2.Font = fontDialog1.Font;
                radioButton3.Font = fontDialog1.Font;
                radioButton4.Font = fontDialog1.Font;
                Returns.Font = fontDialog1.Font;
                listView1.Font = fontDialog1.Font;
                button1.Font = fontDialog1.Font;
                button2.Font = fontDialog1.Font;
                button3.Font = fontDialog1.Font;
                button4.Font = fontDialog1.Font;
                button5.Font = fontDialog1.Font;
                button6.Font = fontDialog1.Font;
                textBox1.Font = fontDialog1.Font;
                button9.Font = fontDialog1.Font;
                button10.Font = fontDialog1.Font;
                Admin.Font = fontDialog1.Font;
            }
        }

        private void allTicketsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            radioButton1.Checked = true;
            radioButton1_CheckedChanged(this, e);
        }

        private void nonAdminToolStripMenuItem_Click(object sender, EventArgs e)
        {
            radioButton3.Checked = true;
            radioButton3_CheckedChanged(this, e);
        }

        private void adminToolStripMenuItem_Click(object sender, EventArgs e)
        {
            radioButton2.Checked = true;
            radioButton2_CheckedChanged(this, e);
        }

        private void returnsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Returns.Checked = true;
            Returns_CheckedChanged(this, e);
        }

        private void closedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            radioButton4.Checked = true;
            radioButton4_CheckedChanged(this, e);
        }

        private void saveTicketsToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            var Savedata = new Savedates();
            Savedata.Show();
        }

        private void launchChatToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            var chats = new chat();
            chats.Show();
        }

        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            RefreshForm(this, e);
        }

        private void newTicketToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewTicket(this, e);
        }

        private void newReturnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button9_Click(this, e);
        }

        private void newAdminToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Admin_Click(sender, e);
        }

        private void updateTicketToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.FocusedItem != null)
            {
                LaunchUpdate(sender, e);
            }
            else
            {
                MessageBox.Show("Please select an item to update");
            }
        }

        private void closeTicketToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseTicket(sender, e);
        }

        private void deleteTicketToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Delete_Click(sender, e);
        }
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button8_Click(sender, e);
        }


        #endregion

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void notifyIcon1_BalloonTipClicked(object sender, EventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            //Show();
            //this.WindowState = FormWindowState.Normal;

            var Moredet = new Form2(this, DomainAdmin);
            Moredet.Show();
        }

        private void saveByTicketToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.FocusedItem != null)
            {
                if (listView1.FocusedItem.Text.Contains("E"))
                {
                    MessageBox.Show("Unable to save return tickets. Sorry for any inconvenience this causes");
                }
                else
                {
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "Spreadsheet|*.csv";
                    save.ShowDialog();

                    TicketsDBDataContext db = new TicketsDBDataContext();
                    var r = from p in db.Ticketers.Where<Ticketer>(x => x.ID2.ToString().Contains(listView1.FocusedItem.Text))
                            select p;

                    string path = save.FileName.ToString();
                    if (path == "")
                    {
                        return;
                    }
                    string delimiter = ",";

                    //string builder
                    StringBuilder sb = new StringBuilder();

                    //header
                    string[] hrow = new string[] { "TicketID", "Ticketor", "Date Ticket Created", "Ticketor Email", "Ticketor Phone number", "Device ID", "Supported by CNS IT", "Assigned to:", "Issue", "Resolver", "Resolver Email", "Resolver Phone Number", "Resolution", "Date Resolved", "Priority", "Issue Type", "Status", "Additional Notes" };
                    sb.AppendLine(string.Join(delimiter, hrow));

                    foreach (var x in r)
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
                        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select an Item to save!", "Invalid Selection",
MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
    public static class FlashWindow
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        [StructLayout(LayoutKind.Sequential)]
        private struct FLASHWINFO
        {
            /// <summary>
                    /// The size of the structure in bytes.
                    /// </summary>
            public uint cbSize;
            /// <summary>
                    /// A Handle to the Window to be Flashed. The window can be either opened or minimized.
                    /// </summary>
            public IntPtr hwnd;
            /// <summary>
                    /// The Flash Status.
                    /// </summary>
            public uint dwFlags;
            /// <summary>
                    /// The number of times to Flash the window.
                    /// </summary>
            public uint uCount;
            /// <summary>
                    /// The rate at which the Window is to be flashed, in milliseconds. If Zero, the function uses the default cursor blink rate.
                    /// </summary>
            public uint dwTimeout;
        }

        /// <summary>
            /// Stop flashing. The system restores the window to its original stae.
            /// </summary>
        public const uint FLASHW_STOP = 0;

        /// <summary>
            /// Flash the window caption.
            /// </summary>
        public const uint FLASHW_CAPTION = 1;

        /// <summary>
            /// Flash the taskbar button.
            /// </summary>
        public const uint FLASHW_TRAY = 2;

        /// <summary>
            /// Flash both the window caption and taskbar button.
            /// This is equivalent to setting the FLASHW_CAPTION | FLASHW_TRAY flags.
            /// </summary>
        public const uint FLASHW_ALL = 3;

        /// <summary>
            /// Flash continuously, until the FLASHW_STOP flag is set.
            /// </summary>
        public const uint FLASHW_TIMER = 4;

        /// <summary>
            /// Flash continuously until the window comes to the foreground.
            /// </summary>
        public const uint FLASHW_TIMERNOFG = 12;


        /// <summary>
            /// Flash the spacified Window (Form) until it recieves focus.
            /// </summary>
            /// <param name="form">The Form (Window) to Flash.</param>
            /// <returns></returns>
        public static bool Flash(System.Windows.Forms.Form form)
        {
            // Make sure we're running under Windows 2000 or later
            if (Win2000OrLater)
            {
                FLASHWINFO fi = Create_FLASHWINFO(form.Handle, FLASHW_ALL | FLASHW_TIMERNOFG, uint.MaxValue, 0);
                return FlashWindowEx(ref fi);
            }
            return false;
        }

        private static FLASHWINFO Create_FLASHWINFO(IntPtr handle, uint flags, uint count, uint timeout)
        {
            FLASHWINFO fi = new FLASHWINFO();
            fi.cbSize = Convert.ToUInt32(Marshal.SizeOf(fi));
            fi.hwnd = handle;
            fi.dwFlags = flags;
            fi.uCount = count;
            fi.dwTimeout = timeout;
            return fi;
        }

        /// <summary>
            /// Flash the specified Window (form) for the specified number of times
            /// </summary>
            /// <param name="form">The Form (Window) to Flash.</param>
            /// <param name="count">The number of times to Flash.</param>
            /// <returns></returns>
        public static bool Flash(System.Windows.Forms.Form form, uint count)
        {
            if (Win2000OrLater)
            {
                FLASHWINFO fi = Create_FLASHWINFO(form.Handle, FLASHW_ALL, count, 0);
                return FlashWindowEx(ref fi);
            }
            return false;
        }

        /// <summary>
            /// Start Flashing the specified Window (form)
            /// </summary>
            /// <param name="form">The Form (Window) to Flash.</param>
            /// <returns></returns>
        public static bool Start(System.Windows.Forms.Form form)
        {
            if (Win2000OrLater)
            {
                FLASHWINFO fi = Create_FLASHWINFO(form.Handle, FLASHW_ALL, uint.MaxValue, 0);
                return FlashWindowEx(ref fi);
            }
            return false;
        }

        /// <summary>
            /// Stop Flashing the specified Window (form)
            /// </summary>
            /// <param name="form"></param>
            /// <returns></returns>
        public static bool Stop(System.Windows.Forms.Form form)
        {
            if (Win2000OrLater)
            {
                FLASHWINFO fi = Create_FLASHWINFO(form.Handle, FLASHW_STOP, uint.MaxValue, 0);
                return FlashWindowEx(ref fi);
            }
            return false;
        }

        /// <summary>
            /// A boolean value indicating whether the application is running on Windows 2000 or later.
            /// </summary>
        private static bool Win2000OrLater
        {
            get { return System.Environment.OSVersion.Version.Major >= 5; }
        }
    }
}
