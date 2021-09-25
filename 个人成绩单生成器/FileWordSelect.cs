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
    public partial class FileWordSelect : Form
    {
        List<string> fileList;
        public FileWordSelect(List<string> fileList)
        {
            InitializeComponent();
            this.fileList = fileList;
        }

        private void FileWordSelect_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < fileList.Count; ++i)
            {
                dataGridView1.Rows.Add();
                int p = 0;
                dataGridView1.Rows[i].Cells[p++].Value = false;
                dataGridView1.Rows[i].Cells[p++].Value = fileList[i];
            }
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
            string path;

            SaveFileDialog saveImageDialog = new SaveFileDialog();
            saveImageDialog.Filter = "Word Files (*.docx;*.doc)|*.docx;*.doc";
            //openImageDialog.Multiselect = false;

            if (saveImageDialog.ShowDialog() == DialogResult.OK)
            {
                //MessageBox.Show(saveImageDialog.FileName);

                path = saveImageDialog.FileName;
            }
            else
            {
                return;
            }

            List<string> fileListSelect = new List<string>();
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                if ((bool)dataGridView1.Rows[i].Cells[0].Value)
                {
                    fileListSelect.Add(fileList[i]);
                }
            }

            if (fileListSelect.Count < 2)
            {
                MessageBox.Show("请选择多个需要合并的文件");
                return;
            }

            string file1 = fileListSelect[fileListSelect.Count - 1];
            fileListSelect.RemoveAt(fileListSelect.Count - 1);
            new WordClass().InsertMerge(file1, fileListSelect , path);

            MessageBox.Show("合并完成");

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
