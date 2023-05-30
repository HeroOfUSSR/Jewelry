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
    public partial class AddOrder : Form
    {
        public int tv2 = 0;

        public List<int> tv = new List<int>();

        Database database = new Database();

        public AddOrder()
        {
            InitializeComponent();
        }

        private void AddOrder_Load(object sender, EventArgs e)
        {
  
            var queryString = $"SELECT id,name FROM jewels";
            var command = new SqlCommand(queryString, database.GetSqlConnection());
            database.openConnection();
            SqlDataReader reader = command.ExecuteReader();
            comboBox1.Items.Clear();
            while (reader.Read())
            {
                comboBox1.Items.Add(reader[1].ToString());
                tv.Add(int.Parse(reader[0].ToString()));
            }
            reader.Close();

        }


        private void button1_Click_1(object sender, EventArgs e)
        {
            database.openConnection();

            var id = textBox1.Text;
            var name = textBox2.Text;
            var kol = textBox4.Text;
            var time = textBox5.Text;


            var a = $"insert into [dbo].[Orders] (id,name,Thing,Kol_vo,time) values('{id}','{name}','{tv2}','{kol}','{time}')";

            var command = new SqlCommand(a, database.GetSqlConnection());

            command.ExecuteNonQuery();

            MessageBox.Show("Запись создана");


            database.CloseConnection();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            tv2 = tv[comboBox1.SelectedIndex];
        }

     
    }
}

