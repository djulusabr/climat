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

        private void DixonGrubbsTable(DataTable table, SQLiteConnection Connect, ClimatData item, string table_name){
            float significance = 0.0f;
            string result = "";
            SQLiteCommand Command = new SQLiteCommand();
            Command.Connection = Connect;
            Command.CommandText = @"SELECT CriticalValue FROM " + table_name + @"
                                    ORDER BY ABS(Asymmetry - @asymmetry),
                                                ABS(SignificanceLevel - @significance),
		                                        ABS(Autocorrelation - @autocorrelation),
		                                        ABS(SampleSize - @samplesize)
                                    LIMIT 1";
            Command.Parameters.AddWithValue("@asymmetry", item.Asymmetry);
            Command.Parameters.AddWithValue("@significance", 5.0);
            Command.Parameters.AddWithValue("@autocorrelation", item.Autocorrelation);
            Command.Parameters.AddWithValue("@samplesize", item.Y.Count);
            SQLiteDataReader sqlReader = Command.ExecuteReader();
            float value = 0f, critval = 0f;
            switch (table_name)
            {
                case "Dixon11": value = item.Dixon1[0]; break;
                case "Dixon1N": value = item.DixonN[0]; break;
                case "Dixon21": value = item.Dixon1[1]; break;
                case "Dixon2N": value = item.DixonN[1]; break;
                case "Dixon31": value = item.Dixon1[2]; break;
                case "Dixon3N": value = item.DixonN[2]; break;
                case "Dixon41": value = item.Dixon1[3]; break;
                case "Dixon4N": value = item.DixonN[3]; break;
                case "Dixon51": value = item.Dixon1[4]; break;
                case "Dixon5N": value = item.DixonN[4]; break;
                case "Grubbs1": value = item.Grubbs1; break;
                case "GrubbsN": value = item.GrubbsN; break;

            }

            if (sqlReader.Read())
                critval = sqlReader.GetFloat(sqlReader.GetOrdinal("CriticalValue"));
            if (critval > value)
                result = "Однороден";
            else
                result = "Неоднороден";

            string maxmin = "";
            string name = "";
            if (table_name[table_name.Length - 1] == 'N')
                maxmin = "max";
            else if (table_name[table_name.Length - 1] == '1')
                maxmin = "min";
            if (table_name.Contains("Dixon"))
                name = "Диксон " + table_name[table_name.Length - 2];
            else if (table_name.Contains("Grubbs"))
                name = "Смирнов-Граббс";

            table.Rows.Add(new object[] { null, maxmin, name, value, critval, significance, result
        });
        }

        public Form2(DataTable table, List<int> stat_list, List<string> month_list)
        {
            InitializeComponent();

            List<ClimatData> data = new List<ClimatData>();
            using (SQLiteConnection Connect = new SQLiteConnection(@"Data Source=ClimatTables.db"))
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
                            IdColumn.AutoIncrement = true;
                            IdColumn.AutoIncrementSeed = 1;
                            IdColumn.AutoIncrementStep = 1;

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
                        
                        grid2.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(grid2_DataBindingComplete);
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
                            IdColumn.AutoIncrement = true;
                            IdColumn.AutoIncrementSeed = 1;
                            IdColumn.AutoIncrementStep = 1;

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
                        DixonGrubbsTable(table3, Connect, item, "Dixon1N");
                        DixonGrubbsTable(table3, Connect, item, "Dixon2N");
                        DixonGrubbsTable(table3, Connect, item, "Dixon3N");
                        DixonGrubbsTable(table3, Connect, item, "Dixon4N");
                        DixonGrubbsTable(table3, Connect, item, "Dixon5N");
                        DixonGrubbsTable(table3, Connect, item, "Dixon11");
                        DixonGrubbsTable(table3, Connect, item, "Dixon21");
                        DixonGrubbsTable(table3, Connect, item, "Dixon31");
                        DixonGrubbsTable(table3, Connect, item, "Dixon41");
                        DixonGrubbsTable(table3, Connect, item, "Dixon51");
                        DixonGrubbsTable(table3, Connect, item, "GrubbsN");
                        DixonGrubbsTable(table3, Connect, item, "Grubbs1");
                        
                        grid3.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(grid3_DataBindingComplete);
                        grid3.DataSource = table3;

                        sub_tpage3.Controls.Add(grid3);
                        month_tcontrol.TabPages.Add(sub_tpage3);

                        stat_tcontrol.TabPages.Add(month_tpage);
                    }
                    tabControl1.TabPages.Add(stat_tpage);
                }
            }
        }

        private void grid2_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            ((DataGridView)sender).Columns["Id"].Visible = false;
            ((DataGridView)sender).Columns["Availability"].HeaderText = "Обеспеченность, P%";
            ((DataGridView)sender).Columns["Value"].HeaderText = "Значение ранжированных осадков, мм";
            ((DataGridView)sender).Columns["Year"].HeaderText = "Год";
            ((DataGridView)sender).AutoResizeColumnHeadersHeight();
            ((DataGridView)sender).AutoResizeColumns();
            ((DataGridView)sender).AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void grid3_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            ((DataGridView)sender).Columns["Id"].Visible = false;
            ((DataGridView)sender).Columns["Extremum"].HeaderText = "Экстремум";
            ((DataGridView)sender).Columns["Criteria"].HeaderText = "Критерий";
            ((DataGridView)sender).Columns["Calculated"].HeaderText = "Расчетное значение";
            ((DataGridView)sender).Columns["Critical"].HeaderText = "Критическое значение";
            ((DataGridView)sender).Columns["Significance"].HeaderText = "Уровень значимости расчетный";
            ((DataGridView)sender).Columns["Result"].HeaderText = "Вывод";
            ((DataGridView)sender).AutoResizeColumnHeadersHeight();
            ((DataGridView)sender).AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
            ((DataGridView)sender).AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            ((DataGridView)sender).Columns["Criteria"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            ((DataGridView)sender).Columns["Result"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}