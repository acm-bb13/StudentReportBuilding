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
    public partial class TheadControl : Form
    {
        List<Student> students;
        int mod = 1;
        string path;
        public TheadControl(int a , int b , int mod , string path , List<Student> students)
        {
            
            InitializeComponent();
            this.students = students;
            this.label1.Text = "第 "+a+" / "+b+" 组    文件生成中，请勿操作！！";
            this.progressBar1.Maximum = b;
            this.progressBar2.Maximum = students.Count * 10;
            this.progressBar1.Value = a - 1;
            this.label2.Text = "正在生成第 1 / "+ students.Count + " 个文件";
            this.mod = mod;
            this.path = path;
        }

        private void TheadControl_Load(object sender, EventArgs e)
        {
            
        }

        public void Begin()
        {
            //遍历
            for (int i = 0; i < students.Count; ++i)
            {
                try
                {
                    if(i == students.Count - 1)
                        progressBar2.Value = progressBar2.Maximum;
                    this.label2.Text = "正在生成第 "+(i + 1)+" / " + students.Count + " 个文件";
                    StudentForm studentForm = new StudentForm(students[i]);
                    studentForm.flush4();
                    studentForm.buildFile(path, mod);
                    studentForm.Close();
                    progressBar2.PerformStep();
                    
                }
                catch (Exception eeee)
                {
                    //异常处理，单个文件生成失败不要影响整体
                }
            }

            this.Close();
        }

    }
}
