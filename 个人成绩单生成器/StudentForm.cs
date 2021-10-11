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
    public partial class StudentForm : Form
    {
        string id;
        string className;
        string sign;
        string oid;
        string name;

        //入学时间
        DateTime dataTime;




        public StudentForm(string id , string className , string sign , string oid , string name , DateTime dataTime)
        {
            InitializeComponent();
            this.id = id;
            this.className = className;
            {
                int temp = int.Parse(oid);
                this.oid = String.Format("{0:D8}", temp);
            }
            this.sign = sign;
            this.name = name;
            this.dataTime = dataTime;

            if (sign == "1")
                this.className = className.Remove(className.Length - 3, 3);
        }

        public StudentForm(Student student)
        {
            InitializeComponent();
            this.id = student.id;
            this.className = student.className;
            {
                int temp = int.Parse(student.oid);
                this.oid = String.Format("{0:D8}", temp);
            }
            this.sign = student.sign;
            this.name = student.name;
            this.dataTime = student.dataTime;

            if (sign == "1")
                this.className = className.Remove(className.Length - 3, 3);
        }

        private void StudentForm_Load(object sender, EventArgs e)
        {
            flush4();
        }



        //void flush3()
        //{
        //    label6.Text = oid;
        //    label7.Text = name;
        //    label8.Text = className;
        //    if (dataGridView1.Rows.Count > 0) { dataGridView1.Rows.Clear(); }
        //    string sql = "select * from testinfo where '" + id + "' = stuId";
        //    MySqlDataReader mySql = SQLManage.GetReader(sql);

        //    //init
        //    if (dataGridView1.Rows.Count > 0) { dataGridView1.Rows.Clear(); }
        //    if (dataGridView1.Columns.Count > 0) { dataGridView1.Columns.Clear(); }
        //    for (int i = 0; i < mySql.FieldCount; i++)
        //    {
        //        DataGridViewTextBoxColumn acCode = new DataGridViewTextBoxColumn();
        //        acCode.Name = mySql.GetName(i);
        //        acCode.DataPropertyName = mySql.GetName(i);
        //        acCode.HeaderText = mySql.GetName(i);
        //        dataGridView1.Columns.Add(acCode);
        //    }

        //    int record = 0;
        //    while (mySql.Read())
        //    {
        //        record++;
        //        int index = dataGridView1.Rows.Add();
        //        int p = 0;
        //        while (mySql.FieldCount > p)
        //        {
        //            dataGridView1.Rows[index].Cells[p].Value = mySql[p++];
        //        }
        //    }
        //    if (record == 0)
        //    {
        //        int index = dataGridView1.Rows.Add();
        //        int p = 0;
        //        dataGridView1.Rows[index].Cells[p++].Value = "暂无数据";
        //        dataGridView1.Rows[index].Cells[p++].Value = "暂无数据";
        //        dataGridView1.Rows[index].Cells[p++].Value = "暂无数据";
        //    }
        //    SQLManage.closeConn();
        //}


        public void flush4()
        {
            label6.Text = oid;
            label7.Text = name;
            label8.Text = className;
            //label10.Text = dataTime.ToString("yyyy-MM-dd");
            if (dataGridView1.Rows.Count > 0) { dataGridView1.Rows.Clear(); }
            string sql = "SELECT courseinfo.name , testinfo.grade , testinfo.xqTime FROM testinfo JOIN courseinfo ON testinfo.courseId = courseinfo.id WHERE testinfo.stuId = '"
                + id + "' ORDER BY xqTime ASC";
            MySqlDataReader mySql = SQLManage.GetReader(sql);

            //初始化表格
            if (dataGridView1.Rows.Count > 0) { dataGridView1.Rows.Clear(); }
            if (dataGridView1.Columns.Count > 0) { dataGridView1.Columns.Clear(); }
            for (int i = 0; i < mySql.FieldCount; i++)
            {
                DataGridViewTextBoxColumn acCode = new DataGridViewTextBoxColumn();
                string t = "";
                if (i == 0) t = "课程名称";
                if (i == 1) t = "成绩";
                if (i == 2) t = "学期";
                acCode.Name = t;
                acCode.DataPropertyName = t;
                acCode.HeaderText = t;
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
                    //if (p == 2)
                    //{
                    //    DateTime t2 = Convert.ToDateTime(mySql[p]);
                    //    int ans = DateTimeSub(dataTime, t2);
                    //    ans = ans / 180 + 1;

                    //    dataGridView1.Rows[index].Cells[p].Value = ans;
                    //    ++p;
                    //    continue;
                    //}
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


        //计算两个datetime的时间差
        int DateTimeSub(DateTime t1, DateTime t2)
        {
            TimeSpan ans = t2.Subtract(t1);
            return ans.Days + 1;
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

        private void button1_Click(object sender, EventArgs e)
        {
            string path = SelectPath();
            if(path != "")
                buildFile(path, 2);
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string path = SelectPath();
            if (path != "")
                buildFile(path, 1);
        }



        //
        public void buildFile(string path , int mod)
        {
            object oMissing = System.Reflection.Missing.Value;
            //创建一个Word应用程序实例  
            Microsoft.Office.Interop.Word._Application oWord = new Microsoft.Office.Interop.Word.Application();
            //设置为不可见  
            oWord.Visible = false;
            //模板文件地址，这里假设在X盘根目录  
            object oTemplate = Application.StartupPath + "\\个人成绩模板-南方冶金学院.docx";  //"D://template.dot";

            if (mod == 1)
            {
                oTemplate = Application.StartupPath + "\\个人成绩模板-南方冶金学院.docx";  //"D://template.dot";
            }

            if(mod == 2)
            {
                oTemplate = Application.StartupPath + "\\个人成绩模板-江西理工大学.docx";  //"D://template.dot";
            }
            
            //以模板为基础生成文档  
            Microsoft.Office.Interop.Word._Document oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);

            string []xqtext = {"第一学期" , "第二学期", "第三学期", "第四学期", "第五学期", "第六学期", "第七学期", "第八学期", "第九学期", "第十学期"};

            try
            {
                int hx = 1 , hc = 2;
                for(int xq = 1; xq <= 10; ++xq)
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                    {
                        int t = int.Parse(dataGridView1.Rows[i].Cells[2].Value.ToString());
                        if(t == xq)
                        {
                            if(hc > 12)
                            {
                                hc = 2;
                                ++hx;
                            }
                            oDoc.Bookmarks.get_Item("bk" + (hx * 2 - 1) + "_" + hc).Range.Text = dataGridView1.Rows[i].Cells[0].Value.ToString();
                            oDoc.Bookmarks.get_Item("bk" + (hx * 2) + "_" + hc).Range.Text = dataGridView1.Rows[i].Cells[1].Value.ToString();
                            oDoc.Bookmarks.get_Item("bk" + (hx * 2 - 1) + "_" + 1).Range.Text = xqtext[xq - 1];
                            ++hc;
                        }
                    }
                    ++hx;
                    hc = 2;
                }


                oDoc.Bookmarks.get_Item("班级").Range.Text = className;
                oDoc.Bookmarks.get_Item("姓名").Range.Text = name;
                oDoc.Bookmarks.get_Item("学号").Range.Text = oid;
            }
            catch (Exception e)
            {

            }


            //弹出保存文件对话框，保存生成的Word  

            object filename = path + "\\" + oid + name + id + ".docx";//"save.doc";//

            oDoc.SaveAs(ref filename, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            oDoc.Close(ref oMissing, ref oMissing, ref oMissing);
            //关闭word  
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            StudentFormSearch studentFormSearch = new StudentFormSearch(name);
            studentFormSearch.ShowDialog();
            StudentForm studentForm = studentFormSearch.studentForm;
            if (studentForm == null) return;
            studentForm.flush4();
            DataGridView dataGridView = studentForm.dataGridView1;
            List<string> classnamelist = new List<string>();
            List<string> goldlist = new List<string>();
            List<string> xqlist = new List<string>();
            List<int> sg = new List<int>();
            for(int i = 0; i < dataGridView.Rows.Count; ++i)
            {
                classnamelist.Add(dataGridView.Rows[i].Cells[0].Value.ToString());
                goldlist.Add(dataGridView.Rows[i].Cells[1].Value.ToString());
                xqlist.Add(dataGridView.Rows[i].Cells[2].Value.ToString());
                sg.Add(1);
            }
            for(int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                for(int index = 0; index < classnamelist.Count; ++index)
                {
                    if (classnamelist[index] == dataGridView1.Rows[i].Cells[0].Value.ToString()
                        && dataGridView1.Rows[i].Cells[2].Value.ToString() == xqlist[index])
                    {
                        sg[index] = 0;
                    }
                    if (classnamelist[index] == dataGridView1.Rows[i].Cells[0].Value.ToString() 
                        && dataGridView1.Rows[i].Cells[1].Value.ToString() == "" && goldlist[index] != ""
                        && dataGridView1.Rows[i].Cells[2].Value.ToString() == xqlist[index]
                        )
                    {
                        dataGridView1.Rows[i].Cells[1].Value = goldlist[index];
                    }
                    if (classnamelist[index] == dataGridView1.Rows[i].Cells[0].Value.ToString()
                        && dataGridView1.Rows[i].Cells[1].Value.ToString() != ""
                        && goldlist[index] != ""
                        && dataGridView1.Rows[i].Cells[1].Value.ToString() != goldlist[index]
                        && dataGridView1.Rows[i].Cells[2].Value.ToString() == xqlist[index])
                    {
                        MessageBox.Show("出现数据冲突，合并终止！！！！\n\n" + "冲突科目为" + classnamelist[index] + "\n" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "与" + goldlist[index] + "不匹配");
                        flush4();
                        return;
                    }
                }
                
            }
            List<string> classnamelist2 = new List<string>();
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                classnamelist2.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
            }
            for (int i = 0; i < dataGridView.Rows.Count; ++i)
            {
                if(sg[i] == 1)
                {
                    int t = dataGridView1.Rows.Add();
                    dataGridView1.Rows[t].Cells[0].Value = dataGridView.Rows[i].Cells[0].Value;
                    dataGridView1.Rows[t].Cells[1].Value = dataGridView.Rows[i].Cells[1].Value;
                    dataGridView1.Rows[t].Cells[2].Value = dataGridView.Rows[i].Cells[2].Value;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            flush4();
        }
    }
}
