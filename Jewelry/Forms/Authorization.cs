using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Jewelry.Forms
{
    public partial class Authorization : Form
    {
        Database database = new Database();
        public Authorization()
        {
            InitializeComponent();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.PasswordChar = '*';
        }


    private void Authorization_Load(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            var login = textBox1.Text;
            var password = textBox2.Text;

            SqlDataAdapter adapter = new SqlDataAdapter();
            DataTable table = new DataTable();

            string query = $"select id,login_user,password_user from register where login_user='{login}' and password_user='{password}'";
            SqlCommand command = new SqlCommand(query, database.GetSqlConnection());

            adapter.SelectCommand = command;
            adapter.Fill(table);

            if (table.Rows.Count == 1)
            {
                MessageBox.Show("Вы успешно авторизировались");

                Authorization auth1 = new Authorization();
                Jewels jews = new Jewels();
                this.Hide();
                jews.ShowDialog();
                this.Show();

            }
            else
            {
                MessageBox.Show("Такого аккаунта не существует");
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Close();
        }
    }
}
