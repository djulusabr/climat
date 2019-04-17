using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using ClimatLibrary;

namespace Climat
{
    public partial class Form2 : Form
    {
        private string StatIdToString(DataTable table, int stat_id)
        {
            foreach (DataRow item in table.Rows)
            {
                if (item.Field<int>(table.Columns["StatId"]) == stat_id)
                {
                    return item.Field<string>(table.Columns["Name"]);
                }
            }
            return "";
        }

        private string MonthToString(string month)
        {
            switch (month)
            {
                case "Jan": return "Январь";
                case "Feb": return "Февраль";
                case "Mar": return "Март";
                case "Apr": return "Апрель";
                case "May": return "Май";
                case "Jun": return "Июнь";
                case "Jul": return "Июль";
                case "Aug": return "Август";
                case "Sep": return "Сентябрь";
                case "Oct": return "Октябрь";
                case "Nov": return "Ноябрь";
                case "Dec": return "Декабрь";
                default: return "";
            }
        }

        private List<float> GetArrayData(DataTable table, int stat_id, string month)
        {
            List<float> list = new List<float>();
            foreach (DataRow item in table.Rows)
            {
                if (item.Field<int>(table.Columns["StatId"]) == stat_id)
                {
                    string str = item.Field<string>(table.Columns[month]);
                    if (!String.IsNullOrEmpty(str))
                    {
                        float res;
                        if (Single.TryParse(str, NumberStyles.Float, CultureInfo.InvariantCulture, out res))
                        {
                            list.Add(res);
                        }
                    }
                }
            }
            list.Sort();
            return list;
        }

        public Form2(DataTable table, List<int> stat_list, List<string> month_list)
        {
            InitializeComponent();

            List<ClimatData> data = new List<ClimatData>();

            richTextBox1.Text = "";

            foreach (int stat_id in stat_list) {
                foreach (string month in month_list) {
                    ClimatData item = new ClimatData(GetArrayData(table, stat_id, month));

                    item = OutlierTest.Dixon(item);
                    item = OutlierTest.Dispersion(item);
                    item = OutlierTest.Grubbs(item);
                    item = OutlierTest.Asymmetry(item);
                    item = OutlierTest.Autocorrelation(item);
                    item = OutlierTest.TDistribution(item);
                    item = OutlierTest.Fisher(item);
                    item = OutlierTest.FDistribution(item);
                    item = OutlierTest.Student(item);

                    richTextBox1.Text += "======================================\n";
                    richTextBox1.Text += "======================================\n";
                    richTextBox1.Text += StatIdToString(table, stat_id) + " - " + MonthToString(month) + "\n\n";

                    richTextBox1.Text += "Критерий Диксона:\n";

                    richTextBox1.Text += "D1n = " + item.DixonN[0] + "\n";
                    richTextBox1.Text += "D2n = " + item.DixonN[1] + "\n";
                    richTextBox1.Text += "D3n = " + item.DixonN[2] + "\n";
                    richTextBox1.Text += "D4n = " + item.DixonN[3] + "\n";
                    richTextBox1.Text += "D5n = " + item.DixonN[4] + "\n";
                    richTextBox1.Text += "D11 = " + item.Dixon1[0] + "\n";
                    richTextBox1.Text += "D21 = " + item.Dixon1[1] + "\n";
                    richTextBox1.Text += "D31 = " + item.Dixon1[2] + "\n";
                    richTextBox1.Text += "D41 = " + item.Dixon1[3] + "\n";
                    richTextBox1.Text += "D51 = " + item.Dixon1[4] + "\n\n";

                    richTextBox1.Text += "Среднее значение = " + item.Y.Average() + "\n";
                    richTextBox1.Text += "Дисперсия = " + item.Dispersion + "\n\n";

                    richTextBox1.Text += "Критерий Смирнова-Граббса:\n";
                    richTextBox1.Text += "Gn = " + item.GrubbsN + "\n";
                    richTextBox1.Text += "G1 = " + item.Grubbs1 + "\n\n";
                    richTextBox1.Text += "Асимметрия = " + item.Asymmetry + "\n";
                    richTextBox1.Text += "Автокорреляция = " + item.Autocorrelation + "\n";
                    richTextBox1.Text += "t-распределение = " + item.TDistribution + "\n\n";

                    richTextBox1.Text += "Критерий Фишера = " + item.Fisher + "\n\n";

                    richTextBox1.Text += "Степень свободы n1 = " + item.FDistribution1 + "\n";
                    richTextBox1.Text += "Степень свободы n2 = " + item.FDistribution2 + "\n\n";

                    richTextBox1.Text += "Критерий Стьюдента = " + item.Student + "\n\n";

                    data.Add(item);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
