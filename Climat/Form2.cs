using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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

        private List<int> GetArrayDataYear(DataTable table, int stat_id, string month)
        {
            Dictionary<int, float> dict = new Dictionary<int, float>();
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
                            dict.Add(item.Field<int>(table.Columns["Year"]), res);
                        }
                    }
                }
            }
            List<int> list = (from entry in dict orderby entry.Value ascending select entry.Key).ToList();
            return list;
        }

        public Form2(DataTable table, List<int> stat_list, List<string> month_list)
        {
            InitializeComponent();

            List<ClimatData> data = new List<ClimatData>();
            using (SQLiteConnection Connect = new SQLiteConnection(@"ClimatTables.db"))
            {
                Connect.Open();
                foreach (int stat_id in stat_list)
                {
                    TabPage stat_tpage = new TabPage(StatIdToString(table, stat_id));
                    TabControl stat_tcontrol = new TabControl();
                    stat_tcontrol.Anchor = (AnchorStyles.Bottom | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left);
                    stat_tcontrol.Dock = DockStyle.Fill;
                    stat_tpage.Controls.Add(stat_tcontrol);

                    foreach (string month in month_list)
                    {
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
                        TabControl month_tcontrol = new TabControl();
                        month_tcontrol.Anchor = (AnchorStyles.Bottom | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left);
                        month_tcontrol.Dock = DockStyle.Fill;
                        month_tpage.Controls.Add(month_tcontrol);

                        TabPage sub_tpage1 = new TabPage("Вычисления");
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

                        sub_tpage1.Controls.Add(rtb);
                        month_tcontrol.TabPages.Add(sub_tpage1);

                        TabPage sub_tpage2 = new TabPage("Ранжированный ряд");
                        DataGridView grid2 = new DataGridView();
                        grid2.Anchor = (AnchorStyles.Bottom | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left);
                        grid2.Dock = DockStyle.Fill;

                        DataTable table2 = new DataTable();
                        {
                            DataColumn IdColumn = new DataColumn("Id", Type.GetType("System.Int32"));
                            IdColumn.Unique = true;
                            IdColumn.AllowDBNull = false;

                            DataColumn AvailabilityColumn = new DataColumn("Availability", Type.GetType("System.Single"));
                            DataColumn ValueColumn = new DataColumn("Value", Type.GetType("System.Single"));
                            DataColumn YearColumn = new DataColumn("Year", Type.GetType("System.Int32"));
                            
                            table2.Columns.Add(IdColumn);
                            table2.Columns.Add(AvailabilityColumn);
                            table2.Columns.Add(ValueColumn);
                            table2.Columns.Add(YearColumn);
                            table2.PrimaryKey = new DataColumn[] { table2.Columns["Id"] };
                        }
                        List<float> list2 = item.Y;
                        List<int> list2b = GetArrayDataYear(table, stat_id, month);
                        list2.Reverse();
                        list2b.Reverse();
                        for (int i = 0; i < list2.Count; i++)
                        {
                            float availability = ((i + 1) / (list2.Count + 1)) * 100; 
                            table2.Rows.Add(new object[] { null, availability, list2[i], list2b[i] });
                        }

                        grid2.DataSource = table2;

                        sub_tpage2.Controls.Add(grid2);
                        month_tcontrol.TabPages.Add(sub_tpage2);

                        TabPage sub_tpage3 = new TabPage("Критерии Диксона и Смирнова-Граббса");
                        DataGridView grid3 = new DataGridView();
                        grid3.Anchor = (AnchorStyles.Bottom | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left);
                        grid3.Dock = DockStyle.Fill;

                        DataTable table3 = new DataTable();
                        {
                            DataColumn IdColumn = new DataColumn("Id", Type.GetType("System.Int32"));
                            IdColumn.Unique = true;
                            IdColumn.AllowDBNull = false;

                            DataColumn ExtremumColumn = new DataColumn("Extremum", Type.GetType("System.String"));
                            DataColumn CriteriaColumn = new DataColumn("Criteria", Type.GetType("System.String"));
                            DataColumn CalculatedColumn = new DataColumn("Calculated", Type.GetType("System.Single"));
                            DataColumn CriticalColumn = new DataColumn("Critical", Type.GetType("System.Single"));
                            DataColumn SignificanceColumn = new DataColumn("Significance", Type.GetType("System.Single"));
                            DataColumn ResultColumn = new DataColumn("Result", Type.GetType("System.String"));

                            table3.Columns.Add(IdColumn);
                            table3.Columns.Add(ExtremumColumn);
                            table3.Columns.Add(CriteriaColumn);
                            table3.Columns.Add(CalculatedColumn);
                            table3.Columns.Add(CriticalColumn);
                            table3.Columns.Add(SignificanceColumn);
                            table3.Columns.Add(ResultColumn);
                            table3.PrimaryKey = new DataColumn[] { table3.Columns["Id"] };
                        }
                        {
                            float significance = 0.0f;
                            string result = "";
                            SQLiteCommand Command = new SQLiteCommand
                            {
                                Connection = Connect,
                                CommandText = @"SELECT * FROM Dixon1N
                                                ORDER BY ABS(Asymmetry - @param1),
                                                         ABS(SignificanceLevel - @param2),
		                                                 ABS(Autocorrelation - @param3),
		                                                 ABS(SampleSize - @param4)
                                                LIMIT 1"
                            };
                            Command.Parameters.Add(new SQLiteParameter("@param1", item.Asymmetry));
                            Command.Parameters.Add(new SQLiteParameter("@param2", 5.0));
                            Command.Parameters.Add(new SQLiteParameter("@param3", item.Autocorrelation));
                            Command.Parameters.Add(new SQLiteParameter("@param4", item.Y.Count));
                            SQLiteDataReader sqlReader = Command.ExecuteReader();
                            float critval = sqlReader.GetFloat(0);
                            if (critval > item.DixonN[0])
                                result = "Однороден";
                            else
                                result = "Неоднороден";

                            table3.Rows.Add(new object[] { null, "max", "Диксон 1", item.DixonN[0], critval, significance, result });
                        }

                        grid3.DataSource = table3;

                        sub_tpage3.Controls.Add(grid3);
                        month_tcontrol.TabPages.Add(sub_tpage3);

                        stat_tcontrol.TabPages.Add(month_tpage);
                    }
                    tabControl1.TabPages.Add(stat_tpage);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}