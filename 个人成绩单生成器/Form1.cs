using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace 个人成绩单生成器
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            flush();
        }

        //读取
        void flush()
        {
            if (dataGridView1.Rows.Count > 0) { dataGridView1.Rows.Clear(); }
            string sql = "select * from classinfo ";

            if (textBox1.Text != "")
            {
                string s = textBox1.Text;
                if (Regex.IsMatch(s, @"^[+-]?\d*[.]?\d*$"))
                {
                    s = Convert.ToInt32(s).ToString();
                }
                sql += "where name like '%" + s + "%' or id like '%" + s + "%' ";
            }

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
            }
            SQLManage.closeConn();
        }

        private void TextBox1_Press(object sender, KeyPressEventArgs e)
        {

            if (e.KeyChar == 13)
            {
                flush();
            }
        }




        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            flush();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int index = dataGridView1.CurrentRow.Index;
            string id = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string name = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string time = dataGridView1.Rows[index].Cells[2].Value.ToString();
            ClassForm classForm = new ClassForm(id, name, time);
            classForm.ShowDialog();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int index = dataGridView1.CurrentRow.Index;
            //label1.Text = dataGridView1.Rows[index].Cells[0].Value.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            StudentSearch studentSearch = new StudentSearch();
            studentSearch.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            flush();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            List<string> classNames = new List<string>();
            List<string> classIds = new List<string>();
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                classIds.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
                classNames.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());
            }
            FileClassSelect fileClassSelect = new FileClassSelect(classNames , classIds);
            fileClassSelect.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            MessageBox.Show("选定目标文件夹后，该程序会读取目标文件夹内所有子文件夹中的.word文件，请尽量不要选取过多的文件避免造成卡顿\n\n并且该功能需要最新版本的Office支持，若无法使用，请将电脑中的Word文档更新至最新版本！！！"
                , "！！！警告！！！", MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = SelectPath();
            if (path == "") return;
            List<string> fileList = fileFun(path);
            FileWordSelect fileWordSelect = new FileWordSelect(fileList);
            fileWordSelect.ShowDialog();

        }



        /*
         * 利用递归将所选文件路径中的全部子文件里的所有'.docx'文件
         * 全部放置到List中。
         */
        List<string> fileFun(string path)
        {
            List<string> fileList = new List<string>();
            string[] fileX = Directory.GetDirectories(path);
            foreach (string s in fileX)
            {
                fileList.AddRange(fileFun(path + "\\" + System.IO.Path.GetFileName(s)));
            }
            fileX = Directory.GetFiles(path);
            for (int i = 0; i < fileX.Length; i++)
            {
                string ss = fileX[i].Remove(0, fileX[i].Length - 4).ToLower();
                if (".doc".Equals(ss) || ".DOC".Equals(ss))
                {
                    fileList.Add(fileX[i]);

                }
                ss = fileX[i].Remove(0, fileX[i].Length - 5).ToLower();
                if (".docx".Equals(ss) || ".DOCX".Equals(ss))
                {
                    fileList.Add(fileX[i]);
                }
            }
            return fileList;
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
