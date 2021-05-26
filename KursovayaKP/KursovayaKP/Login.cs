using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KursovayaKP
{
    public partial class Login : Form
    {
        public Form1 f1;
        public Form2 f2;
        OleDbConnection con = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\Users\The Witcher\Documents\KPKursovaya.accdb");
        public Login()
        {
            InitializeComponent();
            bunifuThinButton21.BackColor = Color.Transparent;
            
        }

        private void bunifuImageButton1_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void bunifuThinButton21_Click(object sender, EventArgs e)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT COUNT(*) FROM Пользователи WHERE Логин = '" + bunifuTextBox1.Text + "' AND Пароль = '" + bunifuTextBox2.Text + "'";
            int a = int.Parse(cmd.ExecuteScalar().ToString());

            if (a == 0)
            {
                MessageBox.Show("Такого пользователя не существует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                f1 = new Form1(this);
                f2 = new Form2(f1);

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT COUNT(*) FROM Пользователи WHERE Логин = '" + bunifuTextBox1.Text + "' AND Пароль = '" + bunifuTextBox2.Text + "' AND ПравоАдмина= 'Да'";
                int x = int.Parse(cmd.ExecuteScalar().ToString());
                if (x != 0)
                {
                    f1.comboBox1.Items.Add("Пользователи");
                    f2.comboBox5.Items.Add("Пользователи");
                }
                this.Hide(); f1.Show();
            }
            con.Close();
        }

        private void bunifuTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void bunifuTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar==13)
            {
                bunifuTextBox2.Focus();
            }
        }

        private void bunifuTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                bunifuThinButton21_Click(sender,e);
            }
        }

        private void bunifuImageButton2_Click(object sender, EventArgs e)
        {
            if (bunifuTextBox2.PasswordChar=='*')
            {
                bunifuTextBox2.PasswordChar = '\0';
            }
            else bunifuTextBox2.PasswordChar = '*';
        }
    }
}

