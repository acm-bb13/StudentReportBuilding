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
    public partial class FileClassSelect : Form
    {

        List<string> classNames, classIds;

        public FileClassSelect(List<string> classNames , List<string> classIds)
        {
            InitializeComponent();
            this.classNames = classNames;
            this.classIds = classIds;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
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

        private void button1_Click(object sender, EventArgs e)
        {
            string path = SelectPath();
            if (path == "") return;
            List<string> classId = new List<string>();
            List<string> className = new List<string>();
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                if ((bool)dataGridView1.Rows[i].Cells[0].Value)
                {
                    classId.Add(classIds[i]);
                    className.Add(classNames[i]);
                }
            }
            if (classId.Count < 1)
            {
                MessageBox.Show("请选择需要生成成绩单的学生");
                return;
            }

            for (int i = 0; i < classId.Count; ++i)
            {
                List<Student> students2;
                
                try
                {
                    students2 = getStudentList(classId[i], className[i]);
                    System.IO.Directory.CreateDirectory(path + "\\" + className[i]);
                    TheadControl theadControl = new TheadControl(i + 1, classId.Count, 2, path + "\\" + className[i], students2);
                    theadControl.Show();
                    theadControl.Begin();

                }
                catch (Exception eee)
                {

                }


            }


            MessageBox.Show("生成完毕！");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string path = SelectPath();
            if (path == "") return;
            List<string> classId = new List<string>();
            List<string> className = new List<string>();
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                if ((bool)dataGridView1.Rows[i].Cells[0].Value)
                {
                    classId.Add(classIds[i]);
                    className.Add(classNames[i]);
                }
            }
            if (classId.Count < 1)
            {
                MessageBox.Show("请选择需要生成成绩单的学生");
                return;
            }

            for(int i = 0; i < classId.Count; ++i)
            {
                List<Student> students2;
                
                try
                {
                    students2 = getStudentList(classId[i], className[i]);
                    System.IO.Directory.CreateDirectory(path + "\\" + className[i]);
                    TheadControl theadControl = new TheadControl(i + 1, classId.Count, 1, path + "\\" + className[i], students2);
                    theadControl.Show();
                    theadControl.Begin();

                }
                catch(Exception eee)
                {

                }

                
            }

            
            MessageBox.Show("生成完毕！");
        }


        //获取指定班级id的所有学生列表
        public List<Student> getStudentList(string id , string name)
        {
            List<Student> students = new List<Student>();
            string sql = "select * from studentinfo where '" + id + "' = iClass";
            MySqlDataReader mySql = SQLManage.GetReader(sql);
            while (mySql.Read())
            {
                Student student = new Student();
                student.id = mySql[0].ToString();
                student.name = mySql[1].ToString();
                student.dataTime = (DateTime)mySql[4];
                student.oid = mySql[7].ToString();
                student.className = name;
                students.Add(student);
            }
            SQLManage.closeConn();
            return students;
        }

        private void FlieClassSelect_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < classNames.Count; ++i)
            {
                dataGridView1.Rows.Add();
                int p = 0;
                dataGridView1.Rows[i].Cells[p++].Value = false;
                dataGridView1.Rows[i].Cells[p++].Value = classNames[i];
                dataGridView1.Rows[i].Cells[p++].Value = classIds[i];
            }
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
    }
}
