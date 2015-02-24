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

namespace Admify
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            if (Properties.Settings.Default.username != "")
            {
                textBox1.Text = Properties.Settings.Default.username;
                this.ActiveControl = textBox2;
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            showAdmifyForm();
        }

        private void showAdmifyForm()
        {
            Properties.Settings.Default.username = textBox1.Text;
            Properties.Settings.Default.Save();
            label1.Visible = false;
            using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
            {
                if (context.ValidateCredentials(textBox1.Text, textBox2.Text))
                {
                    this.Visible = false;
                    Form2 admifyForm = new Form2(textBox1.Text, textBox2.Text, this);
                    admifyForm.Show();              
                } else
                {
                    label1.Visible = true;
                }
            }

        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                showAdmifyForm();
            }
        }
    }
}
