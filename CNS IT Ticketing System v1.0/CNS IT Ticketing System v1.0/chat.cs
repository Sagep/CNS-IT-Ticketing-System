using System;
using System.Windows.Forms;
using System.Data;
using System.Linq;
using System.Drawing;
using System.Windows.Media.Imaging;

namespace CNS_IT_Ticketing_System_v1._0
{
    public partial class chat : Form
    {
        Timer timer = new Timer();
        bool firstload;
        int messageadded;
        int currentmessagecount;
        string username;
        public chat()
        {
            InitializeComponent();
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.chat_FormClosing);
            currentmessagecount = 0;
        }
        private void chat_FormClosing(object sender, FormClosingEventArgs e)
        {
            timer.Stop();
            //gets the datacontext.
            ChatUsersDataContext users = new ChatUsersDataContext();
            var z = from y in users.ConnectedChatusers
                    select y;
            foreach(var x in z)
            {
                if (x.ConnectedChatUser1.Equals(username, StringComparison.InvariantCultureIgnoreCase))
                {
                    users.ConnectedChatusers.DeleteOnSubmit(x);
                }
            }

            try
            {
                users.SubmitChanges();
            }
            catch
            {
                Console.WriteLine("Unable to log out of chat");
                // Provide for exceptions.
            }
        }
        private void loadmessages()
        {
            messageadded = 0;
            //gets the datacontext.
            ChaterDataContext messages = new ChaterDataContext();

            //checks who is logged in. Sets them to be signed in.
            InternalChat test = messages.InternalChats.SingleOrDefault(X => X.Id == 1);
            if (username.Equals("CNS\\Sage.s.porter", StringComparison.InvariantCultureIgnoreCase))
            {
                test.SageConnection = true;
            }
            if (username.Equals("CNS\\Mike.Sisson", StringComparison.InvariantCultureIgnoreCase))
            {
                test.MikeConnection = true;
            }
            messages.SubmitChanges();

            //checks the status of the other user. 
            checkstatus(test);
        }
        private void chat_Load(object sender, EventArgs e)
        {

            username = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            loadmessages();
            textBox1.Select();
            textBox1.Focus();
            ChaterDataContext messages = new ChaterDataContext();
            InternalChat test = messages.InternalChats.SingleOrDefault(X => X.Id == 1);
            checkstatus(test);
            timer.Interval = (10 * 1000);
            //timer.Interval = (20 * 1000);
            timer.Tick += new EventHandler(timer_tick);
            timer.Start();
        }
        private void timer_tick(object sender, EventArgs e)
        {
            Messages.Suspend();
            Messages.Clear();
            loadmessages();
        }

        public static string PlainTextToRtf(string plainText)
        {
            string escapedPlainText = plainText.Replace(@"\", @"\\").Replace("{", @"\{").Replace("}", @"\}");
            string rtf = @"{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Microsoft Sans Serif;}}\viewkind4\uc1\pard\f0\fs17 \b";
            rtf += escapedPlainText.Replace(Environment.NewLine, @" \par ");
            rtf += " }";
            return rtf;
        }

        private void checkstatus(InternalChat test)
        {
            //if (test.MikeConnection == true && test.SageConnection==false)
            //{
            //    label4.Text = "Mike is Connected";
            //    label3.Text = "Sage is not Connected";
            //}
            //if (test.SageConnection == true && test.MikeConnection==false)
            //{
            //    label4.Text = "Mike is not Connected";
            //    label3.Text = "Sage is Connected";
            //}
            //if(test.MikeConnection==true && test.SageConnection==true)
            //{
            //    label3.Text = "Sage is Connected";
            //    label4.Text = "Mike is Connected";
            //}
            
            listView1.Items.Clear();
            bool alreadyentered = false;
            ChatUsersDataContext users = new ChatUsersDataContext();
            var z = from y in users.ConnectedChatusers
                    where y.Id!=0
                    select y;
            foreach(var x in z)
            {
                if(x.ConnectedChatUser1.Equals(username, StringComparison.InvariantCultureIgnoreCase))
                {
                    alreadyentered = true;
                }
                string[] row = { x.ConnectedChatUser1, "" };
                listView1.Items.Add(System.Convert.ToString("")).SubItems.AddRange(row);
            }
            if (!alreadyentered)
            {
                using(ChatUsersDataContext newuser = new ChatUsersDataContext())
                {
                    ConnectedChatuser newusers = new ConnectedChatuser()
                    {
                        ConnectedChatUser1 = username
                    };
                    newuser.ConnectedChatusers.InsertOnSubmit(newusers);
                    newuser.SubmitChanges();
                }
            }

            ChaterDataContext messages = new ChaterDataContext();
            var r = from p in messages.InternalChats
                    where p.Id != 1
                    select p;
            messageadded = 0;
            Messages.Rtf = "";
            foreach (var x in r)
            {

                DateTime dissue = new DateTime();
                dissue = Convert.ToDateTime(x.Datetime);
                DateTime dnow = new DateTime();
                dnow = DateTime.Now;
                TimeSpan duration = dnow - dissue;
                if (duration.TotalHours>1)
                {
                    messages.InternalChats.DeleteOnSubmit(x);
                    messages.SubmitChanges();
                }
                else
                {
                    messageadded += 1;
                    //insertion of string here

                    Messages.SelectedRtf = PlainTextToRtf("\r\n\r\n"+x.Sender);
                    Messages.SelectedRtf = PlainTextToRtf(" "+x.Datetime.ToString()+": ");
                    try
                    {
                        Messages.SelectedRtf = @x.Message;
                    }
                    catch
                    {
                        Messages.SelectedRtf = PlainTextToRtf(x.Message + "\r\n");
                    }
                    if (firstload == false)
                    {
                        currentmessagecount += 1;
                    }

                }
            }
            Messages.SelectionStart = Messages.Text.Length;
            Messages.ScrollToCaret();
            if (messageadded > currentmessagecount)
            {
                FlashWindow.Flash(this);
                currentmessagecount = messageadded;
            }
            firstload = true;
            Messages.Resume();

        }

        private void Send_Click(object sender, EventArgs e)
        {
            using (ChaterDataContext dbContext = new ChaterDataContext())
            {
                var test = new InternalChat
                {
                    Sender = System.Security.Principal.WindowsIdentity.GetCurrent().Name,
                    Datetime = DateTime.Now,
                    Message = textBox1.Rtf
                };
                dbContext.InternalChats.InsertOnSubmit(test);
                dbContext.SubmitChanges();
            }
            textBox1.Text = "";
            Messages.Clear();
            currentmessagecount += 1;
            loadmessages();
            textBox1.Select();
            textBox1.Focus();
        }

        private void Messages_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.LinkText);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Image img = Image.FromFile("\\\\cns-server01\\IT$\\Programs created by Sage\\CNS IT Ticketing System v1.0\\CNS IT Ticketing System v1.0\\ChatThumb.png");
            Clipboard.SetImage(img);
            textBox1.Paste();
            Send_Click(sender, e);
            Clipboard.Clear();
        }
    }
    public static class ControlExtensions
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool LockWindowUpdate(IntPtr hWndLock);

        public static void Suspend(this Control control)
        {
            LockWindowUpdate(control.Handle);
        }

        public static void Resume(this Control control)
        {
            LockWindowUpdate(IntPtr.Zero);
        }

    }

}
