using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace 老数据读取器
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        /*
         * 用于读取单个文件，并将文件的路径放置到DBHelper静态变量中
         * 然后将读取到的数据文件里的所有表名显示在comboBox1供使用
         */
        private void readFile(string fileName)
        {
            //初始化路径点
            DBHelper.route = fileName;
            DBHelper.FileName = System.IO.Path.GetFileName(fileName);
            DBHelper.DirectoryName = System.IO.Path.GetDirectoryName(fileName);

            //清理dataGridView里的数据
            if (dataGridView1.Columns.Count > 0) { dataGridView1.Columns.Clear(); }
            if (dataGridView1.Rows.Count > 0) { dataGridView1.Rows.Clear(); }

            label1.Text = DBHelper.FileName;

            try
            {
                //连接部分
                OleDbConnection myconn = DBHelper.Connection;
                DataTable dt = myconn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                {
                    //将读出的数据存在表格中
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
                }

                //关闭连接
                comboBox1.SelectedIndex = 0;
                DBHelper.connection.Close();
                DBHelper.connection = null;
            }
            catch(Exception ee)
            {
                throw ee;
            }
            
        }


        /*
         * 将读取到的数据库文件里的text表显示到dataGridView1中
         */
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



        /*
         * 将所有数据库的表全部打印到Excel文件中
         */
        private void printX()
        {
            label1.Text = "正在转换"+ DBHelper.FileName+"文件";
            //在对应数据库文件下创建同名文件夹
            string path = DBHelper.route;
            //path = System.IO.Path.GetDirectoryName(path)+"\\转-数据库文件-" + System.IO.Path.GetFileName(path);
            path = path.Remove(path.Length - 4, 4);

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            //然后将所有表一一转换成Excel表格
            for (int i = 1; i <= comboBox1.Items.Count - 1; i++)
            {
                label1.Text = "正在转换" + DBHelper.FileName + "文件"+i+"/"+ (comboBox1.Items.Count - 1);
                comboBox1.SelectedIndex = i;
                showX(comboBox1.Items[i].ToString());
                ExcelTool d = new ExcelTool();
                d.OutputAsExcelFile(dataGridView1, path + "\\" + comboBox1.Items[i].ToString());
            }
            label1.Text = "转换完成";
        }


        /*
         * 尝试读取单个文件
         */
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "数据库文件(*.mdb)|*.mdb|所有文件|*.*";
            ofd.ValidateNames = true;
            ofd.CheckPathExists = true;
            ofd.CheckFileExists = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string strFileName = ofd.FileName;
                try
                {
                    readFile(strFileName);
                }
                catch (Exception eeeee)
                {

                }
                
            }
        }


        /*
         * 将已经读取到的文件显示出来
         */
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
            

        }

        /*
         * 利用递归将所选文件路径中的全部子文件里的所有'.mdb'文件
         * 全部放置到List中。
         */
        static List<string> fileList = new List<string>();
        void fileFun(string path)
        {
            string[] fileX = Directory.GetDirectories(path);
            foreach (string s in fileX)
            {
                fileFun(path + "\\" + System.IO.Path.GetFileName(s));
            }
            fileX = Directory.GetFiles(path);
            for (int i = 0; i < fileX.Length; i++)
            {
                string ss = fileX[i].Remove(0, fileX[i].Length - 4).ToLower();
                if (".mdb".Equals(ss) || ".MDB".Equals(ss))
                {
                    fileList.Add(fileX[i]);

                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            this.Visible = false;
            form2.ShowDialog();
            this.Visible = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if(dataGridView1.Rows.Count <= 0)
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
            MessageBox.Show("转换成功！");
        }

        private void button5_Click(object sender, EventArgs e)
        {

            object oMissing = System.Reflection.Missing.Value;
            //创建一个Word应用程序实例  
            Microsoft.Office.Interop.Word._Application oWord = new Microsoft.Office.Interop.Word.Application();
            //设置为不可见  
            oWord.Visible = false;
            //模板文件地址，这里假设在X盘根目录  
            object oTemplate = Application.StartupPath+ "\\个人成绩模板.docx";  //"D://template.dot";
            //以模板为基础生成文档  
            Microsoft.Office.Interop.Word._Document oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);
            //声明书签数组  
            object[] oBookMark = new object[5];
            //赋值书签名  
            oBookMark[0] = "beizhu";
            oBookMark[1] = "name";
            oBookMark[2] = "sex";
            oBookMark[3] = "birthday";
            oBookMark[4] = "hometown";
            //赋值任意数据到书签的位置  
            oDoc.Bookmarks.get_Item(ref oBookMark[0]).Range.Text = "使用模板实现Word生成";
            oDoc.Bookmarks.get_Item(ref oBookMark[1]).Range.Text = "李四";
            oDoc.Bookmarks.get_Item(ref oBookMark[2]).Range.Text = "女";
            oDoc.Bookmarks.get_Item(ref oBookMark[3]).Range.Text = "1987.06.07";
            oDoc.Bookmarks.get_Item(ref oBookMark[4]).Range.Text = "贺州";
            //弹出保存文件对话框，保存生成的Word  
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Word Document(*.doc)|*.doc";
            sfd.DefaultExt = "Word Document(*.doc)|*.doc";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                object filename = sfd.FileName;//"save.doc";//

                oDoc.SaveAs(ref filename, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);
                oDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                //关闭word  
                oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
            }
        }
    }
}
