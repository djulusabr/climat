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
        ClimatData data;
        public Form2(List<float> _list)
        {
            InitializeComponent();

            data = new ClimatData(_list);

            data = OutlierTest.Dixon(data);
            data = OutlierTest.Dispersion(data);
            data = OutlierTest.Grubbs(data);
            data = OutlierTest.Asymmetry(data);
            data = OutlierTest.Autocorrelation(data);
            data = OutlierTest.TDistribution(data);
            data = OutlierTest.Fisher(data);
            data = OutlierTest.FDistribution(data);
            data = OutlierTest.Student(data);

            richTextBox1.Text = "Критерий Диксона:\n";

            richTextBox1.Text += "D1n = " + data.DixonN[0] + "\n";
            richTextBox1.Text += "D2n = " + data.DixonN[1] + "\n";
            richTextBox1.Text += "D3n = " + data.DixonN[2] + "\n";
            richTextBox1.Text += "D4n = " + data.DixonN[3] + "\n";
            richTextBox1.Text += "D5n = " + data.DixonN[4] + "\n";
            richTextBox1.Text += "D11 = " + data.Dixon1[0] + "\n";
            richTextBox1.Text += "D21 = " + data.Dixon1[1] + "\n";
            richTextBox1.Text += "D31 = " + data.Dixon1[2] + "\n";
            richTextBox1.Text += "D41 = " + data.Dixon1[3] + "\n";
            richTextBox1.Text += "D51 = " + data.Dixon1[4] + "\n\n";

            richTextBox1.Text += "Среднее значение = " + data.Y.Average() + "\n";
            richTextBox1.Text += "Дисперсия = " + data.Dispersion + "\n\n";
            
            richTextBox1.Text += "Критерий Смирнова-Граббса:\n";
            richTextBox1.Text += "Gn = " + data.GrubbsN + "\n";
            richTextBox1.Text += "G1 = " + data.Grubbs1 + "\n\n";
            richTextBox1.Text += "Асимметрия = " + data.Asymmetry + "\n";
            richTextBox1.Text += "Автокорреляция = " + data.Autocorrelation + "\n";
            richTextBox1.Text += "t-распределение = " + data.TDistribution + "\n\n";

            richTextBox1.Text += "Критерий Фишера = " + data.Fisher + "\n\n";

            richTextBox1.Text += "Степень свободы n1 = " + data.FDistribution1 + "\n";
            richTextBox1.Text += "Степень свободы n2 = " + data.FDistribution2 + "\n\n";

            richTextBox1.Text += "Критерий Стьюдента = " + data.Student + "\n\n";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
