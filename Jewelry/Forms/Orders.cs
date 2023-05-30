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
    public partial class Orders : Form
    {
        int selectrow;
        Database database = new Database();
        public Orders()
        {
            InitializeComponent();
        }
        private void CreatColumnsOrders()
        {
            dataGridView1.Columns.Add("id", "id");
            dataGridView1.Columns.Add("name", "Название");
            dataGridView1.Columns.Add("THING", "IDукр");
            dataGridView1.Columns[2].Visible = false;
            dataGridView1.Columns.Add("name2", "Наименование");
            dataGridView1.Columns.Add("Kol_vo", "Количество");
            dataGridView1.Columns.Add("time", "Срок готовности");
            dataGridView1.Columns.Add("Istnew", String.Empty);
        }

        private void ReadSingleRow(DataGridView dataGridK, IDataRecord recordK)
        {
            dataGridK.Rows.Add(recordK.GetInt32(0), recordK.GetString(1), recordK.GetInt32(5), recordK.GetString(6), recordK.GetInt32(3), recordK.GetInt32(4), RowState.ModifiedNew);
        }

        private void RefreshDatagrid()
        {
            dataGridView1.Rows.Clear();
            string query = $"select * from Orders,jewels where Orders.Thing=jewels.id";
            SqlCommand cmd = new SqlCommand(query, database.GetSqlConnection());
            database.openConnection();
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ReadSingleRow(dataGridView1, reader);
            }
            reader.Close();


        }


        private void search(DataGridView dwg)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            search(dataGridView1);
        }


        private void deliterow()
        {
            int index = dataGridView1.CurrentCell.RowIndex;

            dataGridView1.Rows[index].Visible = false;

            if (dataGridView1.Rows[index].Cells[0].Value.ToString() == string.Empty)
            {
                dataGridView1.Rows[index].Cells[5].Value = RowState.Deleted;
                return;
            }

            dataGridView1.Rows[index].Cells[5].Value = RowState.Deleted;

        }
        private void del()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                var rowState = (RowState)dataGridView1.Rows[i].Cells[5].Value;

                if (rowState == RowState.Existed)
                {

                    continue;
                }

                if (rowState == RowState.Deleted)
                {
                    var id = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value);

                    var deletequery = $"delete from Orders where id = {id}";

                    var command = new SqlCommand(deletequery, database.GetSqlConnection());

                    command.ExecuteNonQuery();

                }
            }
        }


        private void button1_Click_1(object sender, EventArgs e)
        {
            AddOrder orders = new AddOrder();
            orders.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sure = "Вы уверенны?";
            string title = "Вы уверенны? ";

            var result = MessageBox.Show(sure, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                deliterow();
                del();
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            int i, j;

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            ExcelApp.Cells[1, 1] = "id";
            ExcelApp.Cells[1, 2] = "Название";
            ExcelApp.Cells[1, 3] = "Цена";
            ExcelApp.Cells[1, 4] = "Вид спорта";
            ExcelApp.Cells[1, 5] = "Количество";


            for (i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (j = 0; j < dataGridView1.ColumnCount; j++)
                {

                    ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;

                }
            }
            //Вызываем нашу созданную эксельку.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            RefreshDatagrid();
        }

        private void Orders_Load_1(object sender, EventArgs e)
        {
            CreatColumnsOrders();
            RefreshDatagrid();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            selectrow = e.RowIndex;
            if (e.RowIndex >= 0)
            {

            }
        }
    }
}

