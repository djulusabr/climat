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

            foreach (int stat_id in stat_list) {
                TabPage stat_tpage = new TabPage(StatIdToString(table, stat_id));

                TabControl stat_tabcontrol = new TabControl();
                stat_tabcontrol.Anchor = (AnchorStyles.Bottom | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left);
                stat_tabcontrol.Dock = DockStyle.Fill;

                stat_tpage.Controls.Add(stat_tabcontrol);

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

                    data.Add(item);

                    TabPage month_tpage = new TabPage(MonthToString(month));
                    RichTextBox rtb = new RichTextBox();
                    rtb.Anchor = (AnchorStyles.Bottom | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left);
                    rtb.Dock = DockStyle.Fill;

                    rtb.Text = "Критерий Диксона:\n";

                    rtb.Text += "D1n = " + item.DixonN[0] + "\n";
                    rtb.Text += "D2n = " + item.DixonN[1] + "\n";
                    rtb.Text += "D3n = " + item.DixonN[2] + "\n";
                    rtb.Text += "D4n = " + item.DixonN[3] + "\n";
                    rtb.Text += "D5n = " + item.DixonN[4] + "\n";
                    rtb.Text += "D11 = " + item.Dixon1[0] + "\n";
                    rtb.Text += "D21 = " + item.Dixon1[1] + "\n";
                    rtb.Text += "D31 = " + item.Dixon1[2] + "\n";
                    rtb.Text += "D41 = " + item.Dixon1[3] + "\n";
                    rtb.Text += "D51 = " + item.Dixon1[4] + "\n\n";

                    rtb.Text += "Среднее значение = " + item.Y.Average() + "\n";
                    rtb.Text += "Дисперсия = " + item.Dispersion + "\n\n";

                    rtb.Text += "Критерий Смирнова-Граббса:\n";
                    rtb.Text += "Gn = " + item.GrubbsN + "\n";
                    rtb.Text += "G1 = " + item.Grubbs1 + "\n\n";
                    rtb.Text += "Асимметрия = " + item.Asymmetry + "\n";
                    rtb.Text += "Автокорреляция = " + item.Autocorrelation + "\n";
                    rtb.Text += "t-распределение = " + item.TDistribution + "\n\n";

                    rtb.Text += "Критерий Фишера = " + item.Fisher + "\n\n";

                    rtb.Text += "Степень свободы n1 = " + item.FDistribution1 + "\n";
                    rtb.Text += "Степень свободы n2 = " + item.FDistribution2 + "\n\n";

                    rtb.Text += "Критерий Стьюдента = " + item.Student + "\n\n";

                    month_tpage.Controls.Add(rtb);

                    stat_tabcontrol.TabPages.Add(month_tpage);
                }
                tabControl1.TabPages.Add(stat_tpage);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
