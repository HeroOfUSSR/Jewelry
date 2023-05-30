using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Jewelry.Forms
{
    enum RowState
    {
        Existed,
        New,
        Modified,
        ModifiedNew,
        Deleted,
    }

    public partial class Jewels : Form
    {
        string dostup;
        int selectrow;
        Database database = new Database();

        public Jewels()
        {
            InitializeComponent();
        }

        public void Dostup(string user)
        {
            dostup = user;
        }



        private void CreatColumns()
        {
            dataGridView1.Columns.Add("id", "id");
            dataGridView1.Columns.Add("name", "Наименование");
            dataGridView1.Columns.Add("price", "Стоимость");
            dataGridView1.Columns.Add("type", "Категория");
            dataGridView1.Columns.Add("kol-vo", "Количество");
            dataGridView1.Columns.Add("Istnew", String.Empty);
        }
        private void ReadSingleRow(DataGridView dataGrid, IDataRecord record)
        {
            dataGrid.Rows.Add(record.GetInt32(0), record.GetString(1), record.GetInt32(2), record.GetString(3), record.GetInt32(4), RowState.ModifiedNew);
        }

        private void RefreshDatagrid(DataGridView dwg)
        {
            dataGridView1.Rows.Clear();
            string query = $"select * from jewels";
            SqlCommand cmd = new SqlCommand(query, database.GetSqlConnection());
            database.openConnection();
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ReadSingleRow(dwg, reader);
            }
            reader.Close();


        }
        private void ClearField()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
        }


        private void Jewels_Load(object sender, EventArgs e)
        {
            CreatColumns();
            RefreshDatagrid(dataGridView1);


            if (dostup == "Пользователь")
            {
                Add.Enabled = false;
                panel1.Enabled = false;
                Del.Enabled = false;
                Red.Enabled = false;
                Sav.Enabled = false;
                button10.Enabled = false;
            }
        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            selectrow = e.RowIndex;
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dataGridView1.Rows[selectrow];
                textBox1.Text = row.Cells[0].Value.ToString();
                textBox2.Text = row.Cells[1].Value.ToString();
                textBox3.Text = row.Cells[2].Value.ToString();
                textBox4.Text = row.Cells[3].Value.ToString();
                textBox5.Text = row.Cells[4].Value.ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {


        }


        private void search(DataGridView dwg)
        {
            dataGridView1.Rows.Clear();

            string search = $"select * from jewels where concat (id, name, price, type, [kol-vo]) like '%" + textBox6.Text + "%'";
            SqlCommand com = new SqlCommand(search, database.GetSqlConnection());
            database.openConnection();
            SqlDataReader reader = com.ExecuteReader();
            while (reader.Read())
            {
                ReadSingleRow(dwg, reader);

            }

            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Такого товара нету");
            }


            reader.Close();
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

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

        private void button5_Click_1(object sender, EventArgs e)
        {
            ClearField();
            RefreshDatagrid(dataGridView1);
        }


        private void update()
        {
            database.openConnection();

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

                    var deletequery = $"delete from jewels where id = {id}";

                    var command = new SqlCommand(deletequery, database.GetSqlConnection());

                    command.ExecuteNonQuery();

                }

                if (rowState == RowState.Modified)
                {
                    var id = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    var name = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    var price = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    var vid = dataGridView1.Rows[i].Cells[3].Value.ToString();
                    var kolvo = dataGridView1.Rows[i].Cells[4].Value.ToString();


                    string query = $"update jewels set name='{name}', price='{price}', type='{vid}', [kol-vo]='{kolvo}' where id='{id}'";

                    var command = new SqlCommand(query, database.GetSqlConnection());

                    command.ExecuteNonQuery();
                }
            }



        }
        private void change()
        {
            var selectedrowsindex = dataGridView1.CurrentCell.RowIndex;

            var id = textBox1.Text;
            var name = textBox2.Text;

            var vid = textBox4.Text;
            var kol = textBox5.Text;

            int price;

            if (dataGridView1.Rows[selectedrowsindex].Cells[0].Value.ToString() != string.Empty)
            {
                if (int.TryParse(textBox3.Text, out price))
                {
                    dataGridView1.Rows[selectedrowsindex].SetValues(id, name, price, vid, kol);
                    dataGridView1.Rows[selectedrowsindex].Cells[5].Value = RowState.Modified;
                }
            }
            else
            {
                MessageBox.Show("Цена должна иметь числовой формат!");
            }
        }


        private void IMP_Click(object sender, EventArgs e)
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

        private void button1_Click_1(object sender, EventArgs e)
        {



        }

        private void button5_Click(object sender, EventArgs e)
        {
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            sortASC(dataGridView1);
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

            sortDESC(dataGridView1);
        }

        private void sortDESC(DataGridView dwg)
        {

            dataGridView1.Rows.Clear();
            string sort = $"SELECT* FROM jewels ORDER BY price DESC";
            SqlCommand cmd = new SqlCommand(sort, database.GetSqlConnection());
            database.openConnection();
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ReadSingleRow(dwg, reader);
            }
            reader.Close();



        }

        private void sortASC(DataGridView dwg)
        {

            dataGridView1.Rows.Clear();
            string sort = $"SELECT* FROM jewels ORDER BY price ASC";
            SqlCommand cmd = new SqlCommand(sort, database.GetSqlConnection());
            database.openConnection();
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ReadSingleRow(dwg, reader);
            }
            reader.Close();



        }

        private void button10_Click(object sender, EventArgs e)
        {
            Orders order = new Orders();
            order.Show();
        }

        private void Add_Click_1(object sender, EventArgs e)
        {
            AddItem additem = new AddItem();
            additem.Show();
        }

        private void Del_Click_1(object sender, EventArgs e)
        {
            string sure = "Вы уверены?";
            string title = "Вы уверены? ";

            var result = MessageBox.Show(sure, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                deliterow();
            }

        }

        private void Red_Click_1(object sender, EventArgs e)
        {
            change();
        }

        private void Sav_Click_1(object sender, EventArgs e)
        {
            update();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            search(dataGridView1);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int i, j;

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
          
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
       
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

            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }


    }
}

