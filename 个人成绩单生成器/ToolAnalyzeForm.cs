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
    public partial class ToolAnalyzeForm : Form
    {
        string fromSql;
        public ToolAnalyzeForm(string fromSql)
        {
            InitializeComponent();
            this.fromSql = fromSql;
        }

        private void ToolAnalyzeForm_Load(object sender, EventArgs e)
        {
            flush2();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            flush2();
        }



        void flush2()
        {
            int sum = 0;
            if (dataGridView1.Rows.Count > 0) { dataGridView1.Rows.Clear(); }
            string sql = "SELECT a.oid , a.`name` , temp.`name` ,  a.xq , a.iClass , a.id , a.inDateTime ,temp.sign FROM( "+fromSql+" )  AS a LEFT  JOIN ( select id,name,sign FROM  classinfo) AS temp ON a.iClass = temp.id  ";


            if (textBox1.Text != "")
            {
                string s = textBox1.Text;
                sql += " where a.name like '%" + s + "%' or a.id like '%" + s + "%' or a.oid like '%" + s + "%' or temp.`name` like '%" + s + "%'";
            }
            sql += " ORDER BY  a.oid ";
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
                if (i == 2)
                {
                    acCode.Name = "班级";
                    acCode.DataPropertyName = "班级";
                    acCode.HeaderText = "班级";
                }
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
                ++sum;
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
            label2.Text = "已搜索出"+sum+"条数据";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count <= 0)
            {
                MessageBox.Show("无数据！\n请先将表格读取出来！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
                return;
            }
            string filePath = "";
            SaveFileDialog s = new SaveFileDialog();
            s.Title = "保存Excel文件";
            s.Filter = "Excel文件(*.xls)|*.xls";
            s.FilterIndex = 1;
            if (s.ShowDialog() == DialogResult.OK)
                filePath = s.FileName;
            else
                return;
            ExcelTool d = new ExcelTool();
            d.OutputAsExcelFile(dataGridView1, filePath.Remove(filePath.Length - 4, 4));
            MessageBox.Show("导出成功！");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            int index = dataGridView1.CurrentRow.Index;
            string id = dataGridView1.Rows[index].Cells[5].Value.ToString();
            string stname = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string oid = dataGridView1.Rows[index].Cells[0].Value.ToString();
            DateTime dateTime = Convert.ToDateTime(dataGridView1.Rows[index].Cells[6].Value.ToString());
            string name = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string sign = dataGridView1.Rows[index].Cells[7].Value.ToString();
            StudentForm studentForm = new StudentForm(id, name, sign, oid, stname, dateTime);
            studentForm.ShowDialog();
            this.Visible = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            List<Student> students = new List<Student>();
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                Student student = new Student();
                student.id = dataGridView1.Rows[i].Cells[5].Value.ToString();
                student.name = dataGridView1.Rows[i].Cells[1].Value.ToString();
                student.dataTime = (DateTime)dataGridView1.Rows[i].Cells[6].Value;
                student.oid = dataGridView1.Rows[i].Cells[0].Value.ToString();
                student.className = dataGridView1.Rows[i].Cells[2].Value.ToString();
                student.sign = dataGridView1.Rows[i].Cells[7].Value.ToString();
                students.Add(student);
            }
            FileStudentsSelect fileStudentsSelect = new FileStudentsSelect(students);
            fileStudentsSelect.ShowDialog();
            this.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            flush2();
        }
    }
}
