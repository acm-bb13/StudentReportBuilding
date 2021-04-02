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

namespace WinFormsApp3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }



        private void showX(string text)
        {
            if ("".Equals(DBHelper.route)) return;
            if ("".Equals(text)) return;
            OleDbConnection myconn = DBHelper.Connection;
            string sql = "SELECT * FROM " + text /*+"班级基本情况"*/;
            OleDbDataReader mySql = DBHelper.GetReader(sql);

            if (dataGridView1.Rows.Count > 0) { dataGridView1.Rows.Clear(); }

            if (dataGridView1.Columns.Count > 0) { dataGridView1.Columns.Clear(); }
            for (int i = 0; i < mySql.FieldCount; i++)
            {
                DataGridViewTextBoxColumn acCode = new DataGridViewTextBoxColumn();
                acCode.Name = mySql.GetName(i);
                acCode.DataPropertyName = mySql.GetName(i);
                acCode.HeaderText = mySql.GetName(i);
                dataGridView1.Columns.Add(acCode);
            }

            int record = 0;
            while (mySql.Read())
            {
                record++;
                int index = dataGridView1.Rows.Add();
                int p = 0;
                while (mySql.FieldCount > p)
                {
                    dataGridView1.Rows[index].Cells[p].Value = mySql[p++];
                }
            }
            if (record == 0)
            {
                int index = dataGridView1.Rows.Add();
                int p = 0;
                dataGridView1.Rows[index].Cells[p++].Value = "暂无数据";
                dataGridView1.Rows[index].Cells[p++].Value = "暂无数据";
                dataGridView1.Rows[index].Cells[p++].Value = "暂无数据";
            }

            DBHelper.connection.Close();
            DBHelper.connection = null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "动作信息文件(*.mdb)|*.mdb|所有文件|*.*";
            ofd.ValidateNames = true;
            ofd.CheckPathExists = true;
            ofd.CheckFileExists = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string strFileName = ofd.FileName;
                DBHelper.route = strFileName;
                if (dataGridView1.Columns.Count > 0) { dataGridView1.Columns.Clear(); }
                if (dataGridView1.Rows.Count > 0) { dataGridView1.Rows.Clear(); }

                label1.Text = System.IO.Path.GetFileName(strFileName);
                OleDbConnection myconn = DBHelper.Connection;
                DataTable dt = myconn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                int n = dt.Rows.Count;
                int m = dt.Columns.IndexOf("TABLE_NAME");
                string[] tableNames = new string[n];
                int q = 0;
                comboBox1.Items.Clear();
                comboBox1.Items.Insert(q++, "请选择");
                for (int i = 0; i < n; i++)
                {
                    DataRow m_DataRow = dt.Rows[i];
                    tableNames[i] = m_DataRow.ItemArray.GetValue(m).ToString();
                    string sql = "select * from " + tableNames[i];
                    comboBox1.Items.Insert(q++, tableNames[i]);
                }
                comboBox1.SelectedIndex = 0;
                DBHelper.connection.Close();
                DBHelper.connection = null;
            }



        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                return;
            }
            if (!"".Equals(DBHelper.route))
            {
                showX(comboBox1.SelectedItem.ToString());
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {
            //label1.Text = DBHelper.route;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelTool d = new ExcelTool();
            d.OutputAsExcelFile(dataGridView1);
        }
    }
}
