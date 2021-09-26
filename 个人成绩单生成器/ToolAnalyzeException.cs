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
    public partial class ToolAnalyzeException : Form
    {
        public ToolAnalyzeException()
        {
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //寻找学号冲突的学生
            string sql = "SELECT studentinfo.oid ,  studentinfo.`name` , studentinfo.xq , studentinfo.iClass , studentinfo.id , studentinfo.inDateTime FROM studentinfo RIGHT OUTER JOIN (SELECT studentinfo.oid   FROM studentinfo  GROUP BY studentinfo.oid having count(studentinfo.oid)>=2) AS temp ON studentinfo.oid = temp.oid";
            ToolAnalyzeForm toolAnalyzeForm = new ToolAnalyzeForm(sql);
            toolAnalyzeForm.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //寻找相同学号相同姓名学生
            string sql = "SELECT studentinfo.oid ,  studentinfo.`name` , studentinfo.xq , studentinfo.iClass , studentinfo.id , studentinfo.inDateTime  FROM studentinfo RIGHT OUTER  JOIN (SELECT studentinfo.oid, studentinfo.`name` FROM studentinfo  GROUP BY studentinfo.oid, studentinfo.`name`  having count(studentinfo.oid) >= 2) AS temp ON studentinfo.`name` = temp.`name` AND studentinfo.oid = temp.oid ORDER BY studentinfo.`name`";
            ToolAnalyzeForm toolAnalyzeForm = new ToolAnalyzeForm(sql);
            toolAnalyzeForm.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sql = "SELECT a.oid , a.`name` , temp.`name` as name2 ,  a.xq , a.iClass , a.id , a.inDateTime ,temp.sign FROM( SELECT studentinfo.oid ,  studentinfo.`name` , studentinfo.xq , studentinfo.iClass , studentinfo.id , studentinfo.inDateTime FROM studentinfo RIGHT OUTER  JOIN ( SELECT studentinfo.`name` FROM studentinfo  GROUP BY studentinfo.`name`  having count(studentinfo.`name`)>=2) AS temp ON studentinfo.`name` = temp.`name` ORDER BY studentinfo.`name`)  AS a LEFT  JOIN ( select id,name,sign FROM  classinfo) AS temp ON a.iClass = temp.id ORDER BY a.`name`";
            ToolAnalyzeForm toolAnalyzeForm = new ToolAnalyzeForm(sql);
            toolAnalyzeForm.ShowDialog();
        }
    }
}
