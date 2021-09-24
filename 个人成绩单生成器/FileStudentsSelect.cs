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
    public partial class FileStudentsSelect : Form
    {
        List<Student> students;
        public FileStudentsSelect(List<Student> students)
        {
            InitializeComponent();
            this.students = students;
        }

        private void FileStudentsSelect_Load(object sender, EventArgs e)
        {
            for(int i = 0; i < students.Count; ++i)
            {
                dataGridView1.Rows.Add();
                int p = 0;
                dataGridView1.Rows[i].Cells[p++].Value = false;
                dataGridView1.Rows[i].Cells[p++].Value = students[i].className;
                dataGridView1.Rows[i].Cells[p++].Value = students[i].name;
                dataGridView1.Rows[i].Cells[p++].Value = students[i].oid;
                dataGridView1.Rows[i].Cells[p++].Value = students[i].id;
                dataGridView1.Rows[i].Cells[p++].Value = students[i].dataTime;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = SelectPath();
            if (path == "") return;
            List<Student> students2 = new List<Student>();
            for(int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                if ((bool)dataGridView1.Rows[i].Cells[0].Value)
                {
                    students2.Add(students[i]);
                }
            }
            if(students2.Count < 1)
            {
                MessageBox.Show("请选择需要生成成绩单的学生");
                return;
            }
            TheadControl theadControl = new TheadControl(1,1,2,path,students2);
            theadControl.Show();
            theadControl.Begin();
            MessageBox.Show("生成完毕！");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string path = SelectPath();
            if (path == "") return;
            List<Student> students2 = new List<Student>();
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                if ((bool)dataGridView1.Rows[i].Cells[0].Value)
                {
                    students2.Add(students[i]);
                }
            }
            if (students2.Count < 1)
            {
                MessageBox.Show("请选择需要生成成绩单的学生");
                return;
            }
            TheadControl theadControl = new TheadControl(1, 1, 1, path, students2);
            theadControl.Show();
            theadControl.Begin();
            MessageBox.Show("生成完毕！");
        }


        //选择文件路径
        private string SelectPath()
        {
            string path = string.Empty;
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = fbd.SelectedPath;
            }
            return path;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for(int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                dataGridView1.Rows[i].Cells[0].Value = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                dataGridView1.Rows[i].Cells[0].Value = false;
            }
        }
    }
}
