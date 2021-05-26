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
using Application = Microsoft.Office.Interop.Excel.Application;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;

namespace KursovayaKP
{
    public partial class Form2 : Form
    {
        private Application xlExcel;
        private Workbook xlWorkBook;
        public Form1 f1;
            OleDbConnection con = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\Users\The Witcher\Documents\KPKursovaya.accdb");
        public Form2(Form1 f1)
        {
            
            InitializeComponent();
            bunifuFlatButton1.Normalcolor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton2.Normalcolor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton3.Normalcolor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton4.Normalcolor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton1.BackColor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton2.BackColor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton3.BackColor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton4.BackColor = System.Drawing.Color.FromArgb(26, 32, 50);
            Login login = new Login();
            Form1 form1 = new Form1(login);
        
        }
        public Login login = new Login();
        
        
        public void loadTable(string selectTable)
        {
            Form1 f1 = new Form1(login);
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = selectTable;
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT COUNT(*) FROM Пользователи WHERE Логин = '" + bunifuTextBox1.Text + "' AND Пароль = '" + bunifuTextBox2.Text + "' AND ПравоАдмина= 'Да'";
            int x = int.Parse(cmd.ExecuteScalar().ToString());
            if (x != 0)
            {
                comboBox5.Items.Add("Пользователи");
            }
            con.Close();
            
            bunifuTextBox1.Visible = false;
            bunifuTextBox2.Visible = false;
            bunifuTextBox3.Visible = false;
            bunifuTextBox4.Visible = false;
            comboBox1.Visible = false;
            comboBox2.Visible = false;
            comboBox3.Visible = false;
            comboBox4.Visible = false;
            bunifuDatePicker2.Visible = false;
            checkBox1.Visible = false;
            comboBox5.DropDownStyle = ComboBoxStyle.DropDownList;
            if (bunifuCustomLabel1.Text == "Область")
            {
                loadTable(Queries.selectOblast);
                dataGridView1.Columns[0].Visible = false;
                bunifuTextBox1.Visible = true;
                bunifuTextBox1.PlaceholderText = "Название";
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            }
            else
            if (bunifuCustomLabel1.Text == "Налоги")
            {
                loadTable(Queries.selectNalog);
                dataGridView1.Columns[0].Visible = false;
                bunifuTextBox1.Visible = true;
                bunifuTextBox2.Visible = true;
                bunifuTextBox1.PlaceholderText = "Наименование";
                bunifuTextBox2.PlaceholderText = "Сумма 1 платежа";
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            else
            if (bunifuCustomLabel1.Text == "Налогоплательщики")
            {
                loadTable(Queries.selectClient);
                dataGridView1.Columns[0].Visible = false;
                string[] array = getOblast().Select(n => n.ToString()).ToArray();
                comboBox1.Items.AddRange(array);
                bunifuTextBox1.Visible = true;
                bunifuTextBox2.Visible = true;
                bunifuTextBox3.Visible = true;
                bunifuTextBox4.Visible = true;
                comboBox1.Visible = true;
                bunifuFlatButton5.Visible = true;
                bunifuTextBox1.PlaceholderText = "Фамилия";
                bunifuTextBox2.PlaceholderText = "Имя";
                bunifuTextBox3.PlaceholderText = "Отчество";
                bunifuTextBox4.PlaceholderText = "Номер телефона";
                bunifuFlatButton4.Visible = true;
                comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            else
            if (bunifuCustomLabel1.Text == "Операции")
            {
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                comboBox3.Items.Clear();
                comboBox4.Items.Clear();
                loadTable(Queries.selectOperacii);
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns[1].HeaderText = "Фамилия";
                dataGridView1.Columns[8].HeaderText = "Сотрудник";
                string[] array = getClient().Select(n => n.ToString()).ToArray();
                comboBox1.Items.AddRange(array);
                string[] array2 = getNalog().Select(n => n.ToString()).ToArray();
                comboBox2.Items.AddRange(array2);
                string[] array3 = getVidOplati().Select(n => n.ToString()).ToArray();
                comboBox3.Items.AddRange(array3);
                string[] array4 = getSotrudniki().Select(n => n.ToString()).ToArray();
                comboBox4.Items.AddRange(array4);
                comboBox1.Visible = true;
                comboBox2.Visible = true;
                comboBox3.Visible = true;
                comboBox4.Visible = true;
                comboBox1.Location = new System.Drawing.Point(667, 40);
                comboBox2.Location = new System.Drawing.Point(667, 71);
                comboBox3.Location = new System.Drawing.Point(667, 102);
                comboBox4.Location = new System.Drawing.Point(667, 133);
                bunifuFlatButton8.Visible = true;
                comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
                comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
                comboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
                comboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
                bunifuDatePicker2.Visible = true;
                checkBox1.Visible = true;
                bunifuDatePicker2.Location = new System.Drawing.Point(667, 164);
                checkBox1.Location = new System.Drawing.Point(667, 195);

                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            else
            if (bunifuCustomLabel1.Text == "ВидОплаты")
            {

                loadTable(Queries.selectVidOplati);
                dataGridView1.Columns[0].Visible = false;
                bunifuTextBox1.Visible = true;
                bunifuTextBox1.PlaceholderText = "Наименование";
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            else
            if (bunifuCustomLabel1.Text == "Банки")
            {

                loadTable(Queries.selectBank);
                dataGridView1.Columns[0].Visible = false;
                bunifuTextBox1.Visible = true;
                bunifuTextBox2.Visible = true;
                bunifuTextBox3.Visible = true;
                bunifuTextBox1.PlaceholderText="Название";
                bunifuTextBox2.PlaceholderText="Адрес";
                bunifuTextBox3.PlaceholderText="Номер телефона";
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            else
            if (bunifuCustomLabel1.Text == "Почты")
            {

                loadTable(Queries.selectPochta);
                dataGridView1.Columns[0].Visible = false;
                bunifuTextBox1.Visible = true;
                bunifuTextBox2.Visible = true;
                bunifuTextBox3.Visible = true;
                bunifuTextBox1.PlaceholderText = "Название";
                bunifuTextBox2.PlaceholderText = "Адрес";
                bunifuTextBox3.PlaceholderText = "Номер телефона";
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            else
            if (bunifuCustomLabel1.Text == "Сотрудники")
            {

                loadTable(Queries.selectSotrudnik);
                dataGridView1.Columns[0].Visible = false;
                bunifuTextBox1.Visible = true;
                bunifuTextBox2.Visible = true;
                bunifuTextBox3.Visible = true;
                bunifuTextBox4.Visible = true;
                bunifuTextBox1.PlaceholderText = "Фамилия";
                bunifuTextBox2.PlaceholderText = "Имя";
                bunifuTextBox3.PlaceholderText = "Отчество";
                bunifuTextBox4.PlaceholderText = "Номер телефона";
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            }
            else if (bunifuCustomLabel1.Text == "Должники")
            {
                loadTable(Queries.selectDolzh);
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            else if (bunifuCustomLabel1.Text == "Пользователи")
            {

                loadTable(Queries.selectPolz);
                bunifuTextBox1.Visible = true;
                bunifuTextBox2.Visible = true;
                bunifuTextBox1.PlaceholderText = "Логин";
                bunifuTextBox2.PlaceholderText = "Пароль";
                comboBox1.Items.Add("Да");
                comboBox1.Items.Add("Нет");
                comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
                comboBox1.Visible = true;
                dataGridView1.Columns[0].Visible = false; 
                comboBox1.Location = new System.Drawing.Point(667, 102);
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
        }

        private void Form2_Shown(object sender, EventArgs e)
        {

        }


        private List<string> getOblast()
        {
            List<string> opers = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Область WHERE deleted=0";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                opers.Add(item[1].ToString());
            }
            return opers;
        }
        private int getIdByOblast(string nameOper)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT КодОбласти FROM Область where Название = '" + nameOper + "' and deleted=0";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private List<string> getClient()
        {
            List<string> opers = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Налогоплательщики where deleted=0";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                opers.Add(item[1].ToString() + " " + item[2].ToString() + " " + item[3].ToString());
            }
            return opers;
        }
        private int getIdByClient(string nameOper)
        {
            string[] a = nameOper.Split(' ');
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT КодНалогоплательщика FROM Налогоплательщики where Фамилия = '" + a[0] + "' and Имя = '" + a[1] + "' and Отчество = '" + a[2] + "' and deleted=0";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private List<string> getNalog()
        {
            List<string> opers = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Налоги where deleted=0";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                opers.Add(item[1].ToString());
            }
            return opers;
        }
        private int getIdByNalog(string nameOper)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT КодНалога FROM Налоги where Наименование = '" + nameOper + "' and deleted=0";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private List<string> getVidOplati()
        {
            List<string> opers = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM ВидОплаты where deleted=0";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                opers.Add(item[1].ToString());
            }
            return opers;
        }
        private int getIdByVidOplati(string nameOper)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT КодВидаОплаты FROM ВидОплаты where ВидОплаты = '" + nameOper + "' and deleted=0";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private List<string> getSotrudniki()
        {
            List<string> opers = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Сотрудники where deleted=0";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                opers.Add(item[1].ToString());
            }
            return opers;
        }
        private int getIdBySotrudniki(string nameOper)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT КодСотрудника FROM Сотрудники where Фамилия = '" + nameOper + "' and deleted=0";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }





        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void bunifuTextBox1_TextChanged(object sender, EventArgs e)
        {


        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            this.Form2_Load(sender, e);
        }

        

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void bunifuTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (bunifuCustomLabel1.Text == "Область")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "Налоги")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != ' ')
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "Налогоплательщики")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "Сотрудники")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "ВидОплаты")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
        }

        private void bunifuTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (bunifuCustomLabel1.Text == "Налоги")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "Налогоплательщики")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "Сотрудники")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
        }

        private void bunifuTextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (bunifuCustomLabel1.Text == "Налогоплательщики")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "Сотрудники")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            if (bunifuCustomLabel1.Text == "Банки")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            if (bunifuCustomLabel1.Text == "Почты")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
        }

        private void bunifuTextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (bunifuCustomLabel1.Text == "Налогоплательщики")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
            }
            if (bunifuCustomLabel1.Text == "Сотрудники")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
            }
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void bunifuTextBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void bunifuFlatButton3_Click(object sender, EventArgs e)
        {
            if (bunifuCustomLabel1.Text == "Банки")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Update Банки SET Deleted = " + 1 + " WHERE КодБанка=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectBank;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            if (bunifuCustomLabel1.Text == "ВидОплаты")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Update ВидОплаты SET Deleted = " + 1 + " WHERE КодВидаОплаты=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectVidOplati;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            if (bunifuCustomLabel1.Text == "Налоги")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Update Налоги SET Deleted = " + 1 + " WHERE КодНалога=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectNalog;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            if (bunifuCustomLabel1.Text == "Налогоплательщики")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Update Налогоплательщики SET Deleted = " + 1 + " WHERE КодНалогоплательщика=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectClient;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            if (bunifuCustomLabel1.Text == "Операции")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Update Операции SET Deleted = " + 1 + " WHERE КодОперации=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectOperacii;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            if (bunifuCustomLabel1.Text == "Область")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Update Область SET Deleted = " + 1 + " WHERE КодОбласти=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectOblast;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else
            if (bunifuCustomLabel1.Text == "Почты")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Update Почты SET Deleted = " + 1 + " WHERE КодОтделенияПочты=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectPochta;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else
            if (bunifuCustomLabel1.Text == "Сотрудники")
            {
                int ID = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Update Сотрудники SET Deleted = " + 1 + " WHERE КодСотрудника=" + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectSotrudnik;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            bunifuTextBox1.Clear();
            bunifuTextBox2.Clear();
            bunifuTextBox3.Clear();
            bunifuTextBox4.Clear();
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
            comboBox5.Text = "";

        }

        private void bunifuImageButton1_Click(object sender, EventArgs e)
        {
            Form2_Load(sender, e);
        }

        private void bunifuFlatButton4_Click(object sender, EventArgs e)
        {
            if (comboBox5.Visible == true)
            {
                comboBox5.Visible = false;
            }
            else
            {
                comboBox5.Visible = true;
                comboBox5.Focus();
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            bunifuTextBox1.Text = "";
            bunifuTextBox2.Text = "";
            bunifuTextBox3.Text = "";
            bunifuTextBox4.Text = "";
            comboBox1.SelectedItem = null;
            comboBox2.SelectedItem = null;
            comboBox3.SelectedItem = null;
            comboBox4.SelectedItem = null;
            checkBox1.Checked = false;
            bunifuCustomLabel1.Text = comboBox5.Items[comboBox5.SelectedIndex].ToString();
            Form2_Load(sender, e);
        }

        private void button_Click(object sender, EventArgs e)
        {

        }

        private void bunifuFlatButton1_Click(object sender, EventArgs e)
        {

            if (bunifuCustomLabel1.Text == "Банки")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Банки (Название, Адрес, НомерТелефона) VALUES('" + bunifuTextBox1.Text + "','" + bunifuTextBox2.Text + "','" + bunifuTextBox3.Text + "')";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectBank;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Refresh();
            }
            else
            if (bunifuCustomLabel1.Text == "ВидОплаты")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO ВидОплаты (ВидОплаты) VALUES('" + bunifuTextBox1.Text + "')";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectVidOplati;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            if (bunifuCustomLabel1.Text == "Налоги")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Налоги (Наименование, Сумма1Платежа) VALUES('" + bunifuTextBox1.Text + "', " + bunifuTextBox2.Text + ")";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectNalog;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            if (bunifuCustomLabel1.Text == "Налогоплательщики")
            {
                int IdOblast = getIdByOblast(comboBox1.Text);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Налогоплательщики (Фамилия, Имя, Отчество, НомерТелефона,КодОбласти) VALUES('" + bunifuTextBox1.Text + "', '" + bunifuTextBox2.Text + "','" + bunifuTextBox3.Text + "', '" + bunifuTextBox4.Text + "', " + IdOblast + ")";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectClient;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            if (bunifuCustomLabel1.Text == "Операции")
            {
                string a = "";
                if (checkBox1.Checked == true)
                {
                    a = "Да";
                }
                else
                {
                    a = "Нет";
                }
                int IdCLient = getIdByClient(comboBox1.Text);
                int IdNalog = getIdByNalog(comboBox2.Text);
                int IdSotrudnik = getIdBySotrudniki(comboBox4.Text);
                int IdVid = getIdByVidOplati(comboBox3.Text);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Операции (КодНалогоплательщика, КодНалога, ДатаОперации, Оплачено, КодВидаОплаты, КодСотрудника) VALUES(" + IdCLient + ", " + IdNalog + ", '" + bunifuDatePicker2.Value.ToString("dd.MM.yyyy") + "', '" + a + "'," + IdVid + ", " + IdSotrudnik + " )";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectOperacii;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            if (bunifuCustomLabel1.Text == "Область")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Область (Название) VALUES( '" + bunifuTextBox1.Text + "')";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectOblast;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            if (bunifuCustomLabel1.Text == "Почты")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Почты (Название, Адрес, НомерТелефона) VALUES( '" + bunifuTextBox1.Text + "', '" + bunifuTextBox2.Text + "', '" + bunifuTextBox3.Text + "')";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectPochta;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            if (bunifuCustomLabel1.Text == "Сотрудники")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Сотрудники (Фамилия, Имя, Отчество, НомерТелефона) VALUES( '" + bunifuTextBox1.Text + "', '" + bunifuTextBox2.Text + "', '" + bunifuTextBox3.Text + "', '" + bunifuTextBox4.Text + "')";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectSotrudnik;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                if (bunifuCustomLabel1.Text == "Пользователи")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Пользователи (Логин,Пароль,ПравоАдмина) VALUES( '" + bunifuTextBox1.Text + "', '" + bunifuTextBox2.Text + "', '" + comboBox1.Text + "')";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectPolz;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);


            }
            bunifuTextBox1.Clear();
            bunifuTextBox2.Clear();
            bunifuTextBox3.Clear();
            bunifuTextBox4.Clear();
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
            comboBox5.Text = "";
            


        }

        private void bunifuImageButton2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void bunifuFlatButton2_Click(object sender, EventArgs e)
        {
            if (bunifuCustomLabel1.Text == "Банки")
            {
                int ID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()); // ЭТО ОЧЕНЬ  ВАЖНАЯ СТРОКА, КТО НЕ  ПРОЧТИАЛ - ТОТ НЕ ПРОЧИТАЛ
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Банки SET Название='" + bunifuTextBox1.Text + "', Адрес='" + bunifuTextBox2.Text + "', НомерТелефона='" + bunifuTextBox3.Text + "' WHERE КодБанка= " + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectBank;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (bunifuCustomLabel1.Text == "ВидОплаты")
            {
                int ID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()); // ЭТО ОЧЕНЬ  ВАЖНАЯ СТРОКА, КТО НЕ  ПРОЧТИАЛ - ТОТ НЕ ПРОЧИТАЛ
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE ВидОплаты SET ВидОплаты='" + bunifuTextBox1.Text + "' WHERE КодВидаОплаты= " + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectVidOplati;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (bunifuCustomLabel1.Text == "Налоги")
            {
                int ID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()); // ЭТО ОЧЕНЬ  ВАЖНАЯ СТРОКА, КТО НЕ  ПРОЧТИАЛ - ТОТ НЕ ПРОЧИТАЛ
                 // ЭТО ОЧЕНЬ  ВАЖНАЯ СТРОКА, КТО НЕ  ПРОЧТИАЛ - ТОТ НЕ ПРОЧИТАЛ
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Налоги SET Наименование='" + bunifuTextBox1.Text + "', Сумма1Платежа= " + bunifuTextBox2.Text + " WHERE КодНалога= " + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectNalog;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (bunifuCustomLabel1.Text == "Налогоплательщики")
            {
                int IdObl = getIdByOblast(comboBox1.Text);
                int ID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()); // ЭТО ОЧЕНЬ  ВАЖНАЯ СТРОКА, КТО НЕ  ПРОЧТИАЛ - ТОТ НЕ ПРОЧИТАЛ
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Налогоплательщики SET Фамилия='" + bunifuTextBox1.Text + "', Имя='" + bunifuTextBox2.Text + "', Отчество='" + bunifuTextBox3.Text + "', НомерТелефона = '" + bunifuTextBox4.Text + "', КодОбласти = " + IdObl + " WHERE КодНалогоплательщика= " + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectClient;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (bunifuCustomLabel1.Text == "Область")
            {
                int ID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()); // ЭТО ОЧЕНЬ  ВАЖНАЯ СТРОКА, КТО НЕ  ПРОЧТИАЛ - ТОТ НЕ ПРОЧИТАЛ
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Область SET Название='" + bunifuTextBox1.Text + "' WHERE КодОбласти= " + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectOblast;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (bunifuCustomLabel1.Text == "Операции")
            {
                string a = "";
                if (checkBox1.Checked == true)
                {
                    a = "Да";
                }
                else
                {
                    a = "Нет";
                }
                int IDVidOplati = getIdByVidOplati(comboBox3.Text);
                int idClient = getIdByClient(comboBox1.Text);
                int idNalog = getIdByNalog(comboBox2.Text);
                int idSotr = getIdBySotrudniki(comboBox4.Text);
                int ID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()); // ЭТО ОЧЕНЬ  ВАЖНАЯ СТРОКА, КТО НЕ  ПРОЧТИАЛ - ТОТ НЕ ПРОЧИТАЛ
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Операции SET КодНалогоплательщика=" + idClient + ", КодНалога=" + idNalog + ", ДатаОперации='" + bunifuDatePicker2.Value.ToString("dd.MM.yyyy") + "', Оплачено = '" + a + "', КодВидаОплаты= " + IDVidOplati + ", КодСотрудника = " + idSotr + " WHERE КодОперации= " + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectOperacii;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (bunifuCustomLabel1.Text == "Почты")
            {
                int ID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()); // ЭТО ОЧЕНЬ  ВАЖНАЯ СТРОКА, КТО НЕ  ПРОЧТИАЛ - ТОТ НЕ ПРОЧИТАЛ
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Почты SET Название='" + bunifuTextBox1.Text + "', Адрес='" + bunifuTextBox2.Text + "', НомерТелефона = '" + bunifuTextBox3.Text + "' WHERE КодОтделенияПочты= " + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectPochta;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (bunifuCustomLabel1.Text == "Сотрудники")
            {
                int ID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()); // ЭТО ОЧЕНЬ  ВАЖНАЯ СТРОКА, КТО НЕ  ПРОЧТИАЛ - ТОТ НЕ ПРОЧИТАЛ
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Сотрудники SET Фамилия='" + bunifuTextBox1.Text + "', Имя='" + bunifuTextBox2.Text + "', Отчество = '" + bunifuTextBox3.Text + "', НомерТелефона = '" + bunifuTextBox4.Text + "' WHERE КодСотрудника= " + ID + "";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectSotrudnik;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            bunifuTextBox1.Clear();
            bunifuTextBox2.Clear();
            bunifuTextBox3.Clear();
            bunifuTextBox4.Clear();
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
            comboBox5.Text = "";

        }

        private void bunifuFlatButton5_Click(object sender, EventArgs e)
        {
            loadTable(Queries.selectDolzh);
        }

        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                if (bunifuCustomLabel1.Text == "Банки")
                {
                    bunifuTextBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    bunifuTextBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    bunifuTextBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();

                }
                if (bunifuCustomLabel1.Text == "ВидОплаты")
                {
                    bunifuTextBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();


                }
                if (bunifuCustomLabel1.Text == "Налоги")
                {
                    bunifuTextBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    bunifuTextBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                }
                if (bunifuCustomLabel1.Text == "Налогоплательщики")
                {
                    bunifuTextBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    bunifuTextBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    bunifuTextBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    bunifuTextBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                    comboBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
                }
                if (bunifuCustomLabel1.Text == "Область")
                {
                    bunifuTextBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                }
                if (bunifuCustomLabel1.Text == "Операции")
                {
                    comboBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() + " " + dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString() + " " + dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    comboBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                    bunifuDatePicker2.Value = DateTime.Parse(dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString());
                    if (dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString() == "Да")
                    {
                        checkBox1.Checked = true;
                    }
                    else
                    {
                        checkBox1.Checked = false;
                    }
                    comboBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
                    comboBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();
                }
                if (bunifuCustomLabel1.Text == "Почты")
                {
                    bunifuTextBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    bunifuTextBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    bunifuTextBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                }
                if (bunifuCustomLabel1.Text == "Сотрудники")
                {
                    bunifuTextBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    bunifuTextBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    bunifuTextBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    bunifuTextBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();

                }


            }
        }

        

        private void bunifuTextBox5_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower();
                    }
                }
            }
            if (bunifuCustomLabel1.Text == "Банки")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT КодБанка, Название, Адрес, НомерТелефона FROM Банки WHERE Deleted = 0 and (Название LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Адрес LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or НомерТелефона LIKE '%" + bunifuTextBox5.Text.ToLower() + "%')";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            else
            if (bunifuCustomLabel1.Text == "ВидОплаты")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT КодВидаОплаты, ВидОплаты FROM ВидОплаты WHERE Deleted = 0 and (ВидОплаты LIKE '%" + bunifuTextBox5.Text.ToLower() + "%') ";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (bunifuCustomLabel1.Text == "Налоги")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT КодНалога, Наименование, Сумма1Платежа  FROM Налоги WHERE Deleted = 0 and (Наименование LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Сумма1Платежа LIKE '%" + bunifuTextBox5.Text.ToLower() + "%') ";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (bunifuCustomLabel1.Text == "Налогоплательщики")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Налогоплательщики.КодНалогоплательщика, Налогоплательщики.Фамилия, Налогоплательщики.Имя, Налогоплательщики.Отчество, Налогоплательщики.НомерТелефона, Область.Название FROM Область INNER JOIN Налогоплательщики ON Область.КодОбласти = Налогоплательщики.КодОбласти WHERE Налогоплательщики.Deleted = 0 and (Фамилия LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Имя LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Отчество LIKE '%" + bunifuTextBox5.Text.ToLower() + "%'or НомерТелефона LIKE '%" + bunifuTextBox5.Text.ToLower() + "%')";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (bunifuCustomLabel1.Text == "Область")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT КодОбласти, Название FROM Область WHERE Deleted = 0 and (Название LIKE '%" + bunifuTextBox5.Text.ToLower() + "%')";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (bunifuCustomLabel1.Text == "Операции")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Операции.КодОперации, Налогоплательщики.Фамилия, Налогоплательщики.Имя, Налогоплательщики.Отчество, Налоги.Наименование, Операции.ДатаОперации, Операции.Оплачено, ВидОплаты.ВидОплаты, Сотрудники.Фамилия FROM Налогоплательщики INNER JOIN(Налоги INNER JOIN (Сотрудники INNER JOIN (ВидОплаты INNER JOIN Операции ON ВидОплаты.КодВидаОплаты = Операции.КодВидаОплаты) ON Сотрудники.КодСотрудника = Операции.КодСотрудника) ON Налоги.КодНалога = Операции.КодНалога) ON Налогоплательщики.КодНалогоплательщика = Операции.КодНалогоплательщика WHERE Операции.Deleted = 0 and (Налогоплательщики.Фамилия LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Налогоплательщики.Имя LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Налогоплательщики.Отчество LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Налоги.Наименование LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or ДатаОперации LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Оплачено LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or ВидОплаты.ВидОплаты LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Сотрудники.Фамилия LIKE '%" + bunifuTextBox5.Text.ToLower() + "%')";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (bunifuCustomLabel1.Text == "Почты")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT КодОтделенияПочты, Название, Адрес, НомерТелефона FROM Почты WHERE Deleted = 0 and (Название LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Адрес LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or НомерТелефона LIKE '%" + bunifuTextBox5.Text.ToLower() + "%') ";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (bunifuCustomLabel1.Text == "Сотрудники")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT КодСотрудника, Фамилия, Имя, Отчество, НомерТелефона FROM Сотрудники WHERE Deleted = 0 and (Фамилия LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Имя LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Отчество LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or НомерТелефона LIKE '%" + bunifuTextBox5.Text.ToLower() + "%')";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            if (bunifuCustomLabel1.Text == "Пользователи")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Пользователи.КодПользователя, Пользователи.Логин, Пользователи.Пароль, Пользователи.ПравоАдмина FROM Пользователи WHERE Deleted =0 and (Логин LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Пароль LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or ПравоАдмина LIKE '%" + bunifuTextBox5.Text.ToLower() + "%')";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
        }

        
        public void Export_Data_To_Word(DataGridView DGV, string filename)
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }

                //table format
                oRange.Text = oTemp;

                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                Type.Missing, Type.Missing, ref ApplyBorders,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style
                oDoc.Application.Selection.Tables[1].set_Style("Сетка таблицы");
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = bunifuCustomLabel1.Text;
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
                oDoc.SaveAs2(filename);
            }
        }
        private void bunifuFlatButton6_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Export_Data_To_Word(dataGridView1, sfd.FileName);
            }
        }
        private void QuitExcel()
        {
            if (this.xlWorkBook != null)
            {
                try
                {
                    this.xlWorkBook.Close();
                    Marshal.ReleaseComObject(this.xlWorkBook);
                }
                catch (COMException)
                {
                }

                this.xlWorkBook = null;
            }

            if (this.xlExcel != null)
            {
                try
                {
                    this.xlExcel.Quit();
                    Marshal.ReleaseComObject(this.xlExcel);
                }
                catch (COMException)
                {
                }

                this.xlExcel = null;
            }
        }
        private void CopyGrid()
        {
            // I'm making this up...
            dataGridView1.MultiSelect = true;
            dataGridView1.SelectAll();

            var data = dataGridView1.GetClipboardContent();

            if (data != null)
            {
                Clipboard.SetDataObject(data, true);
            }
            dataGridView1.MultiSelect =false;
        }
        private void bunifuFlatButton7_Click(object sender, EventArgs e)
        {
            try
            {
                this.QuitExcel();
                this.xlExcel = new Application { Visible = false };
                this.xlWorkBook = this.xlExcel.Workbooks.Add(Missing.Value);

                // Copy contents of grid into clipboard, open new instance of excel, a new workbook and sheet,
                // paste clipboard contents into new sheet.
                this.CopyGrid();

                var xlWorkSheet = (Worksheet)this.xlWorkBook.Worksheets.Item[1];

                try
                {
                    var cr = (Range)xlWorkSheet.Cells[1, 1];

                    try
                    {
                        cr.Select();
                        xlWorkSheet.PasteSpecial(cr, NoHTMLFormatting: true);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(cr);
                    }

                    this.xlWorkBook.SaveAs(Path.Combine(Path.GetTempPath(), "ItemUpdate.xls"), XlFileFormat.xlExcel5);
                }
                finally
                {
                    Marshal.ReleaseComObject(xlWorkSheet);
                }

                MessageBox.Show("File Save Successful", "Information", MessageBoxButtons.OK);

                //// If box is checked, show the exported file. Otherwise quit Excel.
                //if (this.checkBox1.Checked)
                //{
                this.xlExcel.Visible = true;
                //}
                //else
                //{
                // this.QuitExcel();
                //}
            }
            catch (SystemException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            // Set the Selection Mode back to Cell Select to avoid conflict with sorting mode.
            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void bunifuFlatButton8_Click(object sender, EventArgs e)
        {
            bunifuGradientPanel2.Visible = true;
        }

        private void bunifuImageButton3_Click(object sender, EventArgs e)
        {
            string s = bunifuDatePicker1.Value.ToString("dd.MM.yyyy");
            string s2 = bunifuDatePicker3.Value.ToString("dd.MM.yyyy");
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT Операции.КодОперации, Налогоплательщики.Фамилия, Налогоплательщики.Имя, Налогоплательщики.Отчество, Налоги.Наименование, Операции.ДатаОперации, Операции.Оплачено, ВидОплаты.ВидОплаты, Сотрудники.Фамилия FROM Сотрудники INNER JOIN(ВидОплаты INNER JOIN (Налоги INNER JOIN (Налогоплательщики INNER JOIN Операции ON Налогоплательщики.КодНалогоплательщика = Операции.КодНалогоплательщика) ON Налоги.КодНалога = Операции.КодНалога) ON ВидОплаты.КодВидаОплаты = Операции.КодВидаОплаты) ON Сотрудники.КодСотрудника = Операции.КодСотрудника WHERE Операции.Deleted = 0 and ДатаОперации >= @dateFirst and ДатаОперации <= @dateSecond";
            cmd.Parameters.AddWithValue("@dateFirst", s);
            cmd.Parameters.AddWithValue("@dateSecond", s2);
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            bunifuGradientPanel2.Visible = false;
        }

        private void bunifuPictureBox1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Данная автоматизированая систему была создала для того, чтобы упростить работу налоговой службы , которая помогает следить за оплатой налогов. Это достигается тем, что в программе был реализовал ряд основных таблиц: налогоплательщик, область, банки, почты, виды оплаты, налоги, операции и сотрудники, . В программном продукте присутствуют все необходимые компоненты для ведения учёта налогоплательщиков. Благодарим за пользование программой", "Уведомление", MessageBoxButtons.OK);
        }
    }
}
