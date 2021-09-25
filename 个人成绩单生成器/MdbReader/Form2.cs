using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace 老数据读取器
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            Control.CheckForIllegalCrossThreadCalls = false;
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
            catch (Exception ee)
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
            DBHelper.connection = null;
            OleDbConnection myconn = DBHelper.Connection;

            //MessageBox.Show(DBHelper.Connection.ToString());
            

            string sql = "SELECT * FROM " + text /*+"班级基本情况"*/;
            OleDbDataReader mySql = DBHelper.GetReader(sql);

            if (dataGridView1.Columns.Count > 0) { dataGridView1.Columns.Clear(); }
            if (dataGridView1.Rows.Count > 0) { dataGridView1.Rows.Clear(); }

            
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
            label1.Text = "正在转换" + DBHelper.FileName + "文件";
            //在对应数据库文件下创建同名文件夹
            string path = DBHelper.route;
            //path = System.IO.Path.GetDirectoryName(path)+"\\转-数据库文件-" + System.IO.Path.GetFileName(path);
            path = path.Remove(path.Length - 4, 4);

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            //然后将所有表一一转换成Excel表格

            progressBar2.Maximum = comboBox1.Items.Count - 1;
            progressBar2.Value = 0;
            progressBar2.Step = 1;
            for (int i = 1; i <= comboBox1.Items.Count - 1; i++)
            {
                progressBar2.PerformStep();
                label1.Text = "正在转换" + DBHelper.FileName + "文件" + i + "/" + (comboBox1.Items.Count - 1);
                comboBox1.SelectedIndex = i;
                showX(comboBox1.Items[i].ToString());
                ExcelTool d = new ExcelTool();
                d.OutputAsExcelFile(dataGridView1, path + "\\" + comboBox1.Items[i].ToString());
            }
            label1.Text = "转换完成";
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
            MessageBox.Show("选定目标文件夹后，该程序会将目标文件夹内所有子文件夹中的.mdb文件转换为.xls的文件，请尽量不要选取过多的文件避免造成卡顿"
                , "！！！警告！！！", MessageBoxButtons.OK, MessageBoxIcon.Information);
            label2.Text = "读取中";
            FolderBrowserDialog dilog = new FolderBrowserDialog();
            dilog.Description = "请选择文件夹";
            string path = "";
            if (dilog.ShowDialog() == DialogResult.OK || dilog.ShowDialog() == DialogResult.Yes)
            {
                path = dilog.SelectedPath;
                fileFun(path);
                int count = fileList.Count;
                int i = 1;
                progressBar1.Maximum = count;
                progressBar1.Value = 0;
                progressBar1.Step = 1;
                foreach (string s in fileList)
                {
                    label2.Text = "正在处理第" + i + "个文件\n"
                        + "共" + count + "个文件";
                    progressBar1.PerformStep();
                    try
                    {
                        readFile(s);
                        printX();
                        //File.Delete(s);
                    }
                    catch (Exception eeeee)
                    {

                    }
                    i++;
                }

                MessageBox.Show("转换完成", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }


        /*
         * 设置一个新的线程来执行批量转换的操作
         */
        

        private void progressBar2_Click(object sender, EventArgs e)
        {

        }
    }
}
