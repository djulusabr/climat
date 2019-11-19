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
using Excel = Microsoft.Office.Interop.Excel;
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
            list.Reverse();
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
            float significance = 5.0f;
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

        private DataTable table;
        private List<int> stat_list;
        private List<string> month_list;
        private string path;

        public Form2(DataTable table, List<int> stat_list, List<string> month_list, string path)
        {
            InitializeComponent();
            this.table = table;
            this.stat_list = stat_list;
            this.month_list = month_list;
            this.path = path;
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
                        item = OutlierTest.Deviation(item);
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

                        //TabPage sub_tpage1 = new TabPage("Вычисления");
                        //RichTextBox rtb = new RichTextBox();
                        //rtb.Anchor = (AnchorStyles.Bottom | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left);
                        //rtb.Dock = DockStyle.Fill;

                        //rtb.Text = "Критерий Диксона:\n";

                        //rtb.Text += "D1n = " + item.DixonN[0] + "\n";
                        //rtb.Text += "D2n = " + item.DixonN[1] + "\n";
                        //rtb.Text += "D3n = " + item.DixonN[2] + "\n";
                        //rtb.Text += "D4n = " + item.DixonN[3] + "\n";
                        //rtb.Text += "D5n = " + item.DixonN[4] + "\n";
                        //rtb.Text += "D11 = " + item.Dixon1[0] + "\n";
                        //rtb.Text += "D21 = " + item.Dixon1[1] + "\n";
                        //rtb.Text += "D31 = " + item.Dixon1[2] + "\n";
                        //rtb.Text += "D41 = " + item.Dixon1[3] + "\n";
                        //rtb.Text += "D51 = " + item.Dixon1[4] + "\n\n";

                        //rtb.Text += "Среднее значение = " + item.Y.Average() + "\n";
                        //rtb.Text += "Дисперсия = " + item.Deviation + "\n\n";

                        //rtb.Text += "Критерий Смирнова-Граббса:\n";
                        //rtb.Text += "Gn = " + item.GrubbsN + "\n";
                        //rtb.Text += "G1 = " + item.Grubbs1 + "\n\n";
                        //rtb.Text += "Асимметрия = " + item.Asymmetry + "\n";
                        //rtb.Text += "Автокорреляция = " + item.Autocorrelation + "\n";
                        //rtb.Text += "t-распределение = " + item.TDistribution + "\n\n";

                        //rtb.Text += "Критерий Фишера = " + item.Fisher + "\n\n";

                        //rtb.Text += "Степень свободы n1 = " + item.FDistribution1 + "\n";
                        //rtb.Text += "Степень свободы n2 = " + item.FDistribution2 + "\n\n";

                        //rtb.Text += "Критерий Стьюдента = " + item.Student + "\n\n";

                        //sub_tpage1.Controls.Add(rtb);
                        //month_tcontrol.TabPages.Add(sub_tpage1);

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

                        TabPage sub_tpage4 = new TabPage("Критерии Фишера и Стьюдента");
                        DataGridView grid4 = new DataGridView();
                        grid4.Anchor = (AnchorStyles.Bottom | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left);
                        grid4.Dock = DockStyle.Fill;

                        DataTable table4 = new DataTable();
                        {
                            {
                                DataColumn IdColumn = new DataColumn("Id", Type.GetType("System.Int32"));
                                IdColumn.Unique = true;
                                IdColumn.AllowDBNull = false;
                                IdColumn.AutoIncrement = true;
                                IdColumn.AutoIncrementSeed = 1;
                                IdColumn.AutoIncrementStep = 1;
                                
                                DataColumn CriteriaColumn = new DataColumn("Criteria", Type.GetType("System.String"));
                                DataColumn CalculatedColumn = new DataColumn("Calculated", Type.GetType("System.Single"));
                                DataColumn CriticalColumn = new DataColumn("Critical", Type.GetType("System.Single"));
                                DataColumn SignificanceColumn = new DataColumn("Significance", Type.GetType("System.Single"));
                                DataColumn ResultColumn = new DataColumn("Result", Type.GetType("System.String"));

                                table4.Columns.Add(IdColumn);
                                table4.Columns.Add(CriteriaColumn);
                                table4.Columns.Add(CalculatedColumn);
                                table4.Columns.Add(CriticalColumn);
                                table4.Columns.Add(SignificanceColumn);
                                table4.Columns.Add(ResultColumn);
                                table4.PrimaryKey = new DataColumn[] { table4.Columns["Id"] };
                            }
                        }

                        {
                            float significance = 5.0f;
                            string result = "";
                            SQLiteCommand Command = new SQLiteCommand();
                            Command.Connection = Connect;
                            Command.CommandText = @"SELECT CriticalValue FROM Fisher1
                                                    ORDER BY ABS(SampleSize - @samplesize),
                                                             ABS(SignificanceLevel - @significance),
		                                                     ABS(Correlation - @correlation),
		                                                     ABS(Autocorrelation - @autocorrelation)
                                                    LIMIT 1";
                            Command.Parameters.AddWithValue("@samplesize", item.Y1.Count);
                            Command.Parameters.AddWithValue("@significance", 5.0);
                            Command.Parameters.AddWithValue("@correlation", 0.0);
                            Command.Parameters.AddWithValue("@autocorrelation", item.Autocorrelation);
                            SQLiteDataReader sqlReader = Command.ExecuteReader();
                            float critval = 0f;

                            if (sqlReader.Read())
                                critval = sqlReader.GetFloat(sqlReader.GetOrdinal("CriticalValue"));
                            if (critval > item.Fisher)
                                result = "Однороден";
                            else
                                result = "Неоднороден";

                            table4.Rows.Add(new object[] { null, "Критерий Фишера", item.Fisher, critval, significance, result });
                        }

                        {
                            float significance = 5.0f;
                            string result = "";
                            SQLiteCommand Command = new SQLiteCommand();
                            Command.Connection = Connect;
                            Command.CommandText = @"SELECT CriticalValue FROM Student1
                                                    ORDER BY ABS(SampleSize - @samplesize),
                                                             ABS(SignificanceLevel - @significance),
		                                                     ABS(Correlation - @correlation),
		                                                     ABS(Autocorrelation - @autocorrelation)
                                                    LIMIT 1";
                            Command.Parameters.AddWithValue("@samplesize", item.Y1.Count);
                            Command.Parameters.AddWithValue("@significance", 5.0);
                            Command.Parameters.AddWithValue("@correlation", 0.0);
                            Command.Parameters.AddWithValue("@autocorrelation", item.Autocorrelation);
                            SQLiteDataReader sqlReader = Command.ExecuteReader();
                            float critval = 0f;

                            if (sqlReader.Read())
                                critval = sqlReader.GetFloat(sqlReader.GetOrdinal("CriticalValue"));
                            if (critval > item.Student)
                                result = "Однороден";
                            else
                                result = "Неоднороден";

                            table4.Rows.Add(new object[] { null, "Критерий Стьюдента", item.Student, critval, significance, result });
                        }

                        grid4.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(grid4_DataBindingComplete);
                        grid4.DataSource = table4;

                        sub_tpage4.Controls.Add(grid4);
                        month_tcontrol.TabPages.Add(sub_tpage4);

                        stat_tcontrol.TabPages.Add(month_tpage);
                    }

                    TabPage sub_tpage5 = new TabPage("Критерии Фишера и Стьюдента");
                    DataGridView grid5 = new DataGridView();
                    grid5.Anchor = (AnchorStyles.Bottom | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left);
                    grid5.Dock = DockStyle.Fill;

                    DataTable table5 = new DataTable();
                    {
                        {
                            DataColumn IdColumn = new DataColumn("Id", Type.GetType("System.Int32"));
                            IdColumn.Unique = true;
                            IdColumn.AllowDBNull = false;
                            IdColumn.AutoIncrement = true;
                            IdColumn.AutoIncrementSeed = 1;
                            IdColumn.AutoIncrementStep = 1;

                            DataColumn CriteriaColumn = new DataColumn("Criteria", Type.GetType("System.String"));
                            DataColumn CalculatedColumn = new DataColumn("Calculated", Type.GetType("System.Single"));
                            DataColumn CriticalColumn = new DataColumn("Critical", Type.GetType("System.Single"));
                            DataColumn SignificanceColumn = new DataColumn("Significance", Type.GetType("System.Single"));
                            DataColumn ResultColumn = new DataColumn("Result", Type.GetType("System.String"));

                            table5.Columns.Add(IdColumn);
                            table5.Columns.Add(CriteriaColumn);
                            table5.Columns.Add(CalculatedColumn);
                            table5.Columns.Add(CriticalColumn);
                            table5.Columns.Add(SignificanceColumn);
                            table5.Columns.Add(ResultColumn);
                            table5.PrimaryKey = new DataColumn[] { table5.Columns["Id"] };
                        }
                    }

                    {
                        
                    }

                    grid5.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(grid4_DataBindingComplete);
                    grid5.DataSource = table5;

                    sub_tpage5.Controls.Add(grid5);
                    stat_tcontrol.TabPages.Add(sub_tpage5);

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

        private void grid4_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            ((DataGridView)sender).Columns["Id"].Visible = false;
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

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            foreach (TabPage stat_tabPage in tabControl1.TabPages)
            {
                // Станция
                Excel.Workbook workBook = excelApp.Workbooks.Add();
                foreach (TabControl stat_tabControl in stat_tabPage.Controls)
                {
                    for (int i = 0; i < stat_tabControl.TabPages.Count; i++)
                    {
                        // Месяц
                        TabPage month_tabPage = stat_tabControl.TabPages[i];
                        Excel.Sheets sheets = workBook.Sheets;
                        Excel.Worksheet workSheet;
                        if (i == 0)
                        {
                            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
                        }
                        else
                        {
                            workSheet = (Excel.Worksheet)sheets.Add(sheets[sheets.Count], Type.Missing, Type.Missing, Type.Missing);
                        }
                        workSheet.Name = month_tabPage.Text;
                        foreach (Control month_control in month_tabPage.Controls)
                        {
                            if (month_control is TabControl month_tabControl)
                            {
                                // Критерий
                                workSheet.Columns[1].ColumnWidth = 16.57;
                                workSheet.Columns[2].ColumnWidth = 17.86;
                                workSheet.Columns[3].ColumnWidth = 8.14;
                                workSheet.Columns[4].ColumnWidth = 5;
                                workSheet.Columns[5].ColumnWidth = 12.14;
                                workSheet.Columns[6].ColumnWidth = 22;
                                workSheet.Columns[7].ColumnWidth = 15.71;
                                workSheet.Columns[8].ColumnWidth = 14.29;
                                workSheet.Columns[9].ColumnWidth = 15.57;
                                workSheet.Columns[10].ColumnWidth = 15.43;

                                workSheet.Cells[1, 1] = "Ранжированный ряд";
                                workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, 3]].Merge();
                                workSheet.Cells[2, 1] = "Обеспеченность, P%";
                                workSheet.Cells[2, 2] = "Значение ранжированных осадков, мм";
                                workSheet.Cells[2, 3] = "Год";
                                {
                                    var range = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[2, 3]];
                                    range.Font.Bold = true;
                                    range.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                    range.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                    range.Style.WrapText = true;
                                    range.Interior.Color = 6968388;
                                    range.Font.Color = Excel.XlRgbColor.rgbWhite;
                                    var borders = range.Cells.Borders;
                                    borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                                    borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                                    borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                                    borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                                    borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                                    borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                                    borders = null;
                                    range = null;
                                }

                                workSheet.Cells[1, 5] = "Критерии Диксона и Смирнова-Граббса";
                                workSheet.Range[workSheet.Cells[1, 5], workSheet.Cells[1, 10]].Merge();
                                workSheet.Cells[2, 5] = "Экстремум";
                                workSheet.Cells[2, 6] = "Критерий";
                                workSheet.Cells[2, 7] = "Расчетное значение";
                                workSheet.Cells[2, 8] = "Критическое значение";
                                workSheet.Cells[2, 9] = "Уровень значимости расчетный";
                                workSheet.Cells[2, 10] = "Вывод";
                                {
                                    var range = workSheet.Range[workSheet.Cells[1, 5], workSheet.Cells[2, 10]];
                                    range.Font.Bold = true;
                                    range.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                    range.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                    range.Style.WrapText = true;
                                    range.Interior.Color = 6968388;
                                    range.Font.Color = Excel.XlRgbColor.rgbWhite;
                                    var borders = range.Cells.Borders;
                                    borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                                    borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                                    borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                                    borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                                    borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                                    borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                                    borders = null;
                                    range = null;
                                }
                                workSheet.Rows[1].RowHeight = 15;
                                workSheet.Rows[2].RowHeight = 45;

                                workSheet.Cells[16, 6] = "Критерии Фишера и Стьюдента";
                                workSheet.Range[workSheet.Cells[16, 6], workSheet.Cells[16, 10]].Merge();
                                workSheet.Cells[17, 6] = "Критерий";
                                workSheet.Cells[17, 7] = "Расчетное значение";
                                workSheet.Cells[17, 8] = "Критическое значение";
                                workSheet.Cells[17, 9] = "Уровень значимости расчетный"; 
                                workSheet.Cells[17, 10] = "Вывод";
                                workSheet.Range[workSheet.Cells[17, 6],  workSheet.Cells[19, 6]].Merge();
                                workSheet.Range[workSheet.Cells[17, 7],  workSheet.Cells[19, 7]].Merge();
                                workSheet.Range[workSheet.Cells[17, 8],  workSheet.Cells[19, 8]].Merge();
                                workSheet.Range[workSheet.Cells[17, 9],  workSheet.Cells[19, 9]].Merge();
                                workSheet.Range[workSheet.Cells[17, 10], workSheet.Cells[19, 10]].Merge();
                                {
                                    var range = workSheet.Range[workSheet.Cells[16, 6], workSheet.Cells[19, 10]];
                                    range.Font.Bold = true;
                                    range.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                    range.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                    range.Style.WrapText = true;
                                    range.Interior.Color = 6968388;
                                    range.Font.Color = Excel.XlRgbColor.rgbWhite;
                                    var borders = range.Cells.Borders;
                                    borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                                    borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                                    borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                                    borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                                    borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                                    borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                                    borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                                    borders = null;
                                    range = null;
                                }
                                workSheet.Rows[16].RowHeight = 15;
                                workSheet.Rows[17].RowHeight = 15;
                                workSheet.Rows[18].RowHeight = 15;
                                workSheet.Rows[19].RowHeight = 15;

                                foreach (TabPage data_tabPage in month_tabControl.TabPages)
                                {
                                    foreach (Control data_control in data_tabPage.Controls)
                                    {
                                        if (data_control is DataGridView data_gridView)
                                        {
                                            // Месячный датагридвью
                                            DataTable dataTable = (DataTable)data_gridView.DataSource;
                                            if (data_tabPage.Text == "Ранжированный ряд")
                                            {
                                                var borders = workSheet.Range[workSheet.Cells[3, 1], workSheet.Cells[3 + dataTable.Rows.Count - 1, 3]].Cells.Borders;
                                                borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                                                borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                                                borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                                                borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                                                borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                                                borders = null;
                                            }
                                            else if (data_tabPage.Text == "Критерии Диксона и Смирнова-Граббса")
                                            {
                                                var borders = workSheet.Range[workSheet.Cells[3, 5], workSheet.Cells[3 + dataTable.Rows.Count - 1, 10]].Cells.Borders;
                                                borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                                                borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                                                borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                                                borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                                                borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                                                borders = null;
                                            }
                                            else if (data_tabPage.Text == "Критерии Фишера и Стьюдента")
                                            {
                                                var borders = workSheet.Range[workSheet.Cells[20, 6], workSheet.Cells[20 + dataTable.Rows.Count - 1, 10]].Cells.Borders;
                                                borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                                                borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                                                borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                                                borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
                                                borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                                                borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                                                borders = null;
                                            }
                                            for (int x = 0; x < dataTable.Rows.Count; x++)
                                            {
                                                DataRow dataRow = dataTable.Rows[x];
                                                if (data_tabPage.Text == "Ранжированный ряд")
                                                {
                                                    workSheet.Cells[x + 3, 1] = dataRow["Availability"].ToString();
                                                    workSheet.Cells[x + 3, 2] = dataRow["Value"].ToString();
                                                    workSheet.Cells[x + 3, 3] = dataRow["Year"].ToString();
                                                    var range = workSheet.Range[workSheet.Cells[x + 3, 1], workSheet.Cells[x + 3, 3]];
                                                    range.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                                    range.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                                    if (x % 2 == 1)
                                                    {
                                                        range.Interior.Color = 14998742;
                                                    }
                                                    range = null;
                                                }
                                                else if (data_tabPage.Text == "Критерии Диксона и Смирнова-Граббса")
                                                {
                                                    workSheet.Cells[x + 3,  5] = dataRow["Extremum"].ToString();
                                                    workSheet.Cells[x + 3,  6] = dataRow["Criteria"].ToString();
                                                    workSheet.Cells[x + 3,  7] = dataRow["Calculated"].ToString();
                                                    workSheet.Cells[x + 3,  8] = dataRow["Critical"].ToString();
                                                    workSheet.Cells[x + 3,  9] = dataRow["Significance"].ToString();
                                                    workSheet.Cells[x + 3, 10] = dataRow["Result"].ToString();
                                                    var range = workSheet.Range[workSheet.Cells[x + 3, 5], workSheet.Cells[x + 3, 10]];
                                                    range.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                                    range.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                                    if (x % 2 == 1)
                                                    {
                                                        range.Interior.Color = 14998742;
                                                    }
                                                    range = null;
                                                }
                                                else if (data_tabPage.Text == "Критерии Фишера и Стьюдента")
                                                {
                                                    workSheet.Cells[x + 20,  6] = dataRow["Criteria"].ToString();
                                                    workSheet.Cells[x + 20,  7] = dataRow["Calculated"].ToString();
                                                    workSheet.Cells[x + 20,  8] = dataRow["Critical"].ToString();
                                                    workSheet.Cells[x + 20,  9] = dataRow["Significance"].ToString();
                                                    workSheet.Cells[x + 20, 10] = dataRow["Result"].ToString();
                                                    var range = workSheet.Range[workSheet.Cells[x + 20, 6], workSheet.Cells[x + 20, 10]];
                                                    range.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                                    range.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                                    if (x % 2 == 1)
                                                    {
                                                        range.Interior.Color = 14998742;
                                                    }
                                                    range = null;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else if (month_control is DataGridView month_gridView)
                            {
                                // Результативный датагридвью
                                // тут нечего писать пока
                            }
                        }
                    }
                }
                workBook.SaveAs(path + "\\" + stat_tabPage.Text + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                workBook.Close();
            }
            excelApp.Quit();
        }
    }
}