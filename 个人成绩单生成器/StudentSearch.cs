using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 个人成绩单生成器
{
    public partial class StudentSearch : Form
    {
        public StudentSearch()
        {
            InitializeComponent();
            flush2();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            flush2();
        }

        void flush2()
        {
            if (dataGridView1.Rows.Count > 0) { dataGridView1.Rows.Clear(); }
            string sql = "select name , oid , 班级 , id  , inDateTime , sign  from studentinfo LEFT  JOIN ( select id as idssss,name as 班级,sign FROM  classinfo) AS temp ON studentinfo.iClass = temp.idssss  ";
            if (textBox1.Text != "")
            {
                string s = textBox1.Text;
                //if (Regex.IsMatch(s, @"^[+-]?\d*[.]?\d*$"))
                //{
                //    s = Convert.ToInt32(s).ToString();
                //}
                sql += " where studentinfo.name like '%" + s + "%' or studentinfo.id like '%" + s + "%' or studentinfo.oid like '%" + s + "%' or 班级 like '%" + s + "%'  ";
            }
            sql += " limit 100 ";

            MySqlDataReader mySql = SQLManage.GetReader(sql);

            //init
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
            SQLManage.closeConn();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            int index = dataGridView1.CurrentRow.Index;
            string id = dataGridView1.Rows[index].Cells[3].Value.ToString();
            string stname = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string oid = dataGridView1.Rows[index].Cells[1].Value.ToString();
            DateTime dateTime = Convert.ToDateTime(dataGridView1.Rows[index].Cells[4].Value.ToString());
            string name = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string sign = dataGridView1.Rows[index].Cells[5].Value.ToString();
            StudentForm studentForm = new StudentForm(id, name , sign , oid , stname, dateTime);
            studentForm.ShowDialog();
            this.Visible = true;
        }

        private void TextBox1_Press(object sender, KeyPressEventArgs e)
        {

            if (e.KeyChar == 13)
            {
                flush2();
            }
        }

        private void StudentSearch_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            flush2();
        }
    }
}
