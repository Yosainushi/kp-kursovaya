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
    
    public partial class Form1 : Form
    {
        public Login login;
        OleDbConnection con = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\Users\The Witcher\Documents\KPKursovaya.accdb");
        public Form1(Login login)
        {
            InitializeComponent();
        }
        
    private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        }

       

        private void Form1_Leave(object sender, EventArgs e)
        {
            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void bunifuImageButton1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("Выберите таблицу", "Ошибка", MessageBoxButtons.OK);
            }
            else
            {
                Login login = new Login();
                Form1 f1 = new Form1(login);
                Form2 f2 = new Form2(f1);
                foreach (var item in comboBox1.Items)
                {
                    if (item.ToString() == "Пользователи")
                    {
                        f2.comboBox5.Items.Add("Пользователи");
                    }
                }
                f2.bunifuCustomLabel1.Text = comboBox1.SelectedItem.ToString();
                this.Hide();
                f2.Show();
            }
        }
        private void bunifuImageButton2_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                bunifuImageButton1_Click(sender, e);
            }
        }
    }
}
