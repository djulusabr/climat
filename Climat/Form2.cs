using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClimatLibrary;

namespace Climat
{
    public partial class Form2 : Form
    {
        List<float> list;
        public Form2(List<float> _list)
        {
            InitializeComponent();
            this.list = _list;

            List<float> dixon = ClimatLibrary.OutlierTest.Dixon(list);
            List<float> smirnovGrubbs = ClimatLibrary.OutlierTest.SmirnovGrubbs(list);

            richTextBox1.Text = "Критерий Диксона:\n";

            richTextBox1.Text += "D1n = " + dixon[0] + "\n";
            richTextBox1.Text += "D2n = " + dixon[1] + "\n";
            richTextBox1.Text += "D3n = " + dixon[2] + "\n";
            richTextBox1.Text += "D4n = " + dixon[3] + "\n";
            richTextBox1.Text += "D5n = " + dixon[4] + "\n";
            richTextBox1.Text += "D11 = " + dixon[5] + "\n";
            richTextBox1.Text += "D21 = " + dixon[6] + "\n";
            richTextBox1.Text += "D31 = " + dixon[7] + "\n";
            richTextBox1.Text += "D41 = " + dixon[8] + "\n";
            richTextBox1.Text += "D51 = " + dixon[9] + "\n\n";

            richTextBox1.Text += "Критерий Смирнова-Граббса:\n";
            richTextBox1.Text += "Среднее значение = " + smirnovGrubbs[0] + "\n";
            richTextBox1.Text += "Дисперсия = " + smirnovGrubbs[1] + "\n";
            richTextBox1.Text += "Gn = " + smirnovGrubbs[2] + "\n";
            richTextBox1.Text += "G1 = " + smirnovGrubbs[3] + "\n";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
