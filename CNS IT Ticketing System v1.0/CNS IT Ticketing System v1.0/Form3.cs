using System;
using System.Windows.Forms;

namespace CNS_IT_Ticketing_System_v1._0
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        //loading of form and settings
        private void Form3_Load(object sender, EventArgs e)
        {
            textBox1.SelectionStart = textBox1.Text.Length;
            ContextMenu blankContextMenu = new ContextMenu();
            textBox1.ContextMenu = blankContextMenu;
        }

        //Shortcut commands
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Escape))
            {
                this.Close();
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
