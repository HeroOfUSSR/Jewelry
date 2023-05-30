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
    public partial class AddItem : Form
    {
        Database database = new Database();

        public AddItem()
        {
            InitializeComponent();
        }
      

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            database.openConnection();

            var id = textBox1.Text;
            var name = textBox2.Text;
            var vid = textBox4.Text;
            var kol = textBox5.Text;
            int price;

            if (int.TryParse(textBox3.Text, out price))
            {
                var a = $"insert into jewels (id,name,price,type,[kol-vo]) values('{id}','{name}','{price}','{vid}','{kol}')";

                var command = new SqlCommand(a, database.GetSqlConnection());

                command.ExecuteNonQuery();

                MessageBox.Show("Запись создана");
            }
            else
            {
                MessageBox.Show("Запись не создана");
            }
            database.CloseConnection();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Close();
        }
    }
}
