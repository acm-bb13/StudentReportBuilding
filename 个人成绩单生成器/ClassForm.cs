using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 个人成绩单生成器
{
    public partial class ClassForm : Form
    {
        //存储读取的信息
        string id, name, time;

        public ClassForm(string id , string name , string time)
        {
            InitializeComponent();
            this.id = id; this.name = name; this.time = time;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            flush2();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            int index = dataGridView1.CurrentRow.Index;
            string id = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string stname = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string oid = dataGridView1.Rows[index].Cells[7].Value.ToString();
            DateTime dateTime = Convert.ToDateTime(dataGridView1.Rows[index].Cells[4].Value.ToString());
            StudentForm studentForm = new StudentForm(id , name , oid , stname , dateTime);
            studentForm.ShowDialog();
            this.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            List<Student> students = new List<Student>();
            for(int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                Student student = new Student();
                student.id = dataGridView1.Rows[i].Cells[0].Value.ToString();
                student.name = dataGridView1.Rows[i].Cells[1].Value.ToString();
                student.dataTime = (DateTime)dataGridView1.Rows[i].Cells[4].Value;
                student.oid = dataGridView1.Rows[i].Cells[7].Value.ToString();
                student.className = name;
                students.Add(student);
            }
            FileStudentsSelect fileStudentsSelect = new FileStudentsSelect(students);
            fileStudentsSelect.ShowDialog();
            this.Visible = true;
        }

        private void ClassForm_Load(object sender, EventArgs e)
        {
            flush2();
            this.Text = name;
        }


        void flush2()
        {
            if (dataGridView1.Rows.Count > 0) { dataGridView1.Rows.Clear(); }
            string sql = "select * from studentinfo where '"+id+"' = iClass";
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
    }
}
