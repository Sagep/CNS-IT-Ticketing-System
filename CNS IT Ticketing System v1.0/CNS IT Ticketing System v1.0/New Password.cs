using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.DirectoryServices.AccountManagement;

namespace CNS_IT_Ticketing_System_v1._0
{
    public partial class NewPassword : Form
    {
        string account;
        public NewPassword(string resetme)
        {
            InitializeComponent();
            account = resetme;
        }

        public void button1_Click(object sender, EventArgs e)
        {
            string final="";
            string initial = textBox1.Text;
            string confirm = textBox2.Text;

            if (initial.Equals(confirm) && initial!="")
                final = initial;

            using (var context = new PrincipalContext(ContextType.Domain))
            {
                using (var user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, account))
                {
                   user.SetPassword(final);
                   user.ExpirePasswordNow();
                   user.UnlockAccount();
                }
            }
            MessageBox.Show("Reset user: " + account+"\r\nPassword: "+final);
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void NewPassword_Load(object sender, EventArgs e)
        {
            MessageBox.Show("Notice: This will reset the password. If you wish to cancel, please click on ''Cancel''");
        }
    }
}
