﻿using System;
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
    public partial class InitForm : Form
    {
        public InitForm()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            this.Visible = false;
            form1.ShowDialog();
            this.Close();
        }
    }
}
