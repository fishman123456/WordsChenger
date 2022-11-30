using Microsoft.Office.Interop.Word;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordsChenger
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            var helper = new WordHelper("Испытания электродвигателя переменного тока напряжением до 1 кВ_.doc");
            var items = new Dictionary<string, string>
            {
                {"<first>", textBox1.Text},
                {"<second>",textBox2.Text},
                {"<three>",textBox3.Text },
                {"<four>",textBox4.Text },
                {"<five>",textBox5.Text },
                {"<six>",textBox6.Text },
                {"<seven>",textBox7.Text },
                {"<eight>",textBox8.Text },
                {"<nine>",textBox9.Text },
                {"<ten>",textBox10.Text },
                {"<eleven>",textBox11.Text },
                {"<twelve>",textBox12.Text },
                {"<thirteen>",textBox13.Text },
                {"<fourteen>",textBox14.Text }
            };
            helper.Process(items);
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

      
    }
}
