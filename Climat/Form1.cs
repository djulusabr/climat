using System;
using System.Data;
using System.Linq;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace Climat
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, Int32 wMsg, bool wParam, Int32 lParam);
        private const int WM_SETREDRAW = 11;
        DataSet climatSet;
        DataTable statListTable, fldTable, wrTable;
        public Form1()
        {
            InitializeComponent();
            climatSet = new DataSet("ClimatSet");
            statListTable = new DataTable("StatListTable");
            fldTable = new DataTable("FldTable");
            wrTable = new DataTable("WrTable");
            climatSet.Tables.Add(statListTable);
            climatSet.Tables.Add(fldTable);
            climatSet.Tables.Add(wrTable);

            {
                DataColumn IdColumn = new DataColumn("Id", Type.GetType("System.Int32"));
                IdColumn.Unique = true;
                IdColumn.AllowDBNull = false;

                DataColumn NameColumn = new DataColumn("Name", Type.GetType("System.String"));

                statListTable.Columns.Add(IdColumn);
                statListTable.Columns.Add(NameColumn);
                statListTable.PrimaryKey = new DataColumn[] { statListTable.Columns["Id"] };
            }

            {
                DataColumn IdColumn = new DataColumn("Id", Type.GetType("System.Int32"));
                IdColumn.Unique = true;
                IdColumn.AllowDBNull = false;

                DataColumn ColNumColumn = new DataColumn("ColNum", Type.GetType("System.Int32"));

                DataColumn FormattingColumn = new DataColumn("Formatting", Type.GetType("System.String"));

                DataColumn ColNameColumn = new DataColumn("ColName", Type.GetType("System.String"));

                fldTable.Columns.Add(IdColumn);
                fldTable.Columns.Add(ColNumColumn);
                fldTable.Columns.Add(FormattingColumn);
                fldTable.Columns.Add(ColNameColumn);
                fldTable.PrimaryKey = new DataColumn[] { fldTable.Columns["Id"] };
            }

            {
                DataColumn IdColumn = new DataColumn("Id", Type.GetType("System.Int32"));
                IdColumn.Unique = true;
                IdColumn.AllowDBNull = false;
                IdColumn.AutoIncrement = true;
                IdColumn.AutoIncrementSeed = 1;
                IdColumn.AutoIncrementStep = 1;

                DataColumn StatIdColumn = new DataColumn("StatId", Type.GetType("System.Int32"));

                DataColumn NameColumn = new DataColumn("Name", Type.GetType("System.String"));

                DataColumn YearColumn = new DataColumn("Year", Type.GetType("System.Int32"));

                DataColumn JanColumn = new DataColumn("Jan", Type.GetType("System.String"));
                DataColumn FebColumn = new DataColumn("Feb", Type.GetType("System.String"));
                DataColumn MarColumn = new DataColumn("Mar", Type.GetType("System.String"));
                DataColumn AprColumn = new DataColumn("Apr", Type.GetType("System.String"));
                DataColumn MayColumn = new DataColumn("May", Type.GetType("System.String"));
                DataColumn JunColumn = new DataColumn("Jun", Type.GetType("System.String"));
                DataColumn JulColumn = new DataColumn("Jul", Type.GetType("System.String"));
                DataColumn AugColumn = new DataColumn("Aug", Type.GetType("System.String"));
                DataColumn SepColumn = new DataColumn("Sep", Type.GetType("System.String"));
                DataColumn OctColumn = new DataColumn("Oct", Type.GetType("System.String"));
                DataColumn NovColumn = new DataColumn("Nov", Type.GetType("System.String"));
                DataColumn DecColumn = new DataColumn("Dec", Type.GetType("System.String"));

                wrTable.Columns.Add(IdColumn);
                wrTable.Columns.Add(StatIdColumn);
                wrTable.Columns.Add(NameColumn);
                wrTable.Columns.Add(YearColumn);
                wrTable.Columns.Add(JanColumn);
                wrTable.Columns.Add(FebColumn);
                wrTable.Columns.Add(MarColumn);
                wrTable.Columns.Add(AprColumn);
                wrTable.Columns.Add(MayColumn);
                wrTable.Columns.Add(JunColumn);
                wrTable.Columns.Add(JulColumn);
                wrTable.Columns.Add(AugColumn);
                wrTable.Columns.Add(SepColumn);
                wrTable.Columns.Add(OctColumn);
                wrTable.Columns.Add(NovColumn);
                wrTable.Columns.Add(DecColumn);
                wrTable.PrimaryKey = new DataColumn[] { wrTable.Columns["Id"] };
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Multiselect = false;
            dlg.Filter = "Archives (*   .rar;*.zip;*.7z)|*.rar;*.zip;*.7z";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (var archive = ZipFile.OpenRead(dlg.FileName))
                    {
                        string fld = "", statlist = "", wr = "";
                        foreach (var entry in archive.Entries)
                        {
                            if (entry.FullName.StartsWith("fld"))
                                fld = entry.FullName;
                            else if (entry.FullName.StartsWith("statlist"))
                                statlist = entry.FullName;
                            else if (entry.FullName.StartsWith("wr"))
                                wr = entry.FullName;
                        }
                        if (statlist.Length > 0 & wr.Length > 0)
                        {
                            string directoryPath = dlg.FileName.Substring(0, dlg.FileName.LastIndexOf('.'));
                            try
                            {
                                ZipFile.ExtractToDirectory(dlg.FileName, directoryPath);
                            }
                            catch (IOException)
                            {
                                //
                            }
                            ReadFromDirectory(directoryPath);
                        }
                        else
                        {
                            MessageBox.Show("Не обнаружены нужные файлы в выбранном архиве", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (InvalidDataException)
                {
                    MessageBox.Show("Выбранный файл не является архивом.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var dlg = new FolderSelectDialog();

            if (dlg.ShowDialog(IntPtr.Zero))
            {
                ReadFromDirectory(dlg.FileName);
            }
        }

        private void ParseWrTable(string line)
        {
            int statid = Convert.ToInt32(line.Substring(0, 5));
            string name = statListTable.Select("Id = " + statid.ToString()).First()["Name"].ToString();
            int year = Convert.ToInt32(line.Substring(6, 4));

            List<string> month_list = GetMonthList();
            int i = 11;
            string jan = "", feb = "", mar = "", apr = "",
                   may = "", jun = "", jul = "", aug = "",
                   sep = "", oct = "", nov = "", dec = "";

            if (month_list.Contains("Jan"))
            {
                jan = line.Substring(i, 5).Replace(" ", string.Empty);
                i += 6;
            }
            if (month_list.Contains("Feb"))
            {
                feb = line.Substring(i, 5).Replace(" ", string.Empty);
                i += 6;
            }
            if (month_list.Contains("Mar"))
            {
                mar = line.Substring(i, 5).Replace(" ", string.Empty);
                i += 6;
            }
            if (month_list.Contains("Apr"))
            {
                apr = line.Substring(i, 5).Replace(" ", string.Empty);
                i += 6;
            }
            if (month_list.Contains("May"))
            {
                may = line.Substring(i, 5).Replace(" ", string.Empty);
                i += 6;
            }
            if (month_list.Contains("Jun"))
            {
                jun = line.Substring(i, 5).Replace(" ", string.Empty);
                i += 6;
            }
            if (month_list.Contains("Jul"))
            {
                jul = line.Substring(i, 5).Replace(" ", string.Empty);
                i += 6;
            }
            if (month_list.Contains("Aug"))
            {
                aug = line.Substring(i, 5).Replace(" ", string.Empty);
                i += 6;
            }
            if (month_list.Contains("Sep"))
            {
                sep = line.Substring(i, 5).Replace(" ", string.Empty);
                i += 6;
            }
            if (month_list.Contains("Oct"))
            {
                oct = line.Substring(i, 5).Replace(" ", string.Empty);
                i += 6;
            }
            if (month_list.Contains("Nov"))
            {
                nov = line.Substring(i, 5).Replace(" ", string.Empty);
                i += 6;
            }
            if (month_list.Contains("Dec"))
            {
                dec = line.Substring(i, 5).Replace(" ", string.Empty);
                i += 6;
            }
            wrTable.Rows.Add(new object[] { null, statid, name, year, jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec });
        }

        private void ReadFromDirectory(string path)
        {
            var fileList = Directory.GetFiles(path, "*.txt");
            string fld = "", statlist = "", wr = "";
            foreach (string str in fileList)
            {
                var fileName = str.Substring(str.LastIndexOf('\\') + 1);
                if (fileName.StartsWith("fld"))
                    fld = str;
                else if (fileName.StartsWith("statlist"))
                    statlist = str;
                else if (fileName.StartsWith("wr"))
                    wr = str;
            }
            if (statlist.Length > 0 && fld.Length > 0)
            {
                statListTable.Clear();
                fldTable.Clear();
                wrTable.Clear();
                try
                {
                    using (StreamReader sr = new StreamReader(statlist, Encoding.GetEncoding(1251)))
                    {
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            int id = Convert.ToInt32(line.Substring(0, 5));
                            string name = line.Substring(6);
                            statListTable.Rows.Add(new object[] { id, name });
                        }
                    }
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Файл не мог быть прочитан:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                try
                {
                    using (StreamReader sr = new StreamReader(fld, Encoding.GetEncoding(1251)))
                    {
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            int id = Convert.ToInt32(line.Substring(0, 2));
                            int colNum = Convert.ToInt32(line.Substring(3, 2));
                            string formatting = line.Substring(6, 7).Replace(" ", string.Empty);
                            string colName = line.Substring(15);
                            fldTable.Rows.Add(new object[] { id, colNum, formatting, colName });
                        }
                    }
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Файл не мог быть прочитан:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                if (wr.Length > 0)
                {
                    try
                    {
                        using (StreamReader sr = new StreamReader(wr, Encoding.GetEncoding(1251)))
                        {
                            string line;
                            while ((line = sr.ReadLine()) != null)
                            {
                                ParseWrTable(line);
                            }
                        }
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show("Файл не мог быть прочитан:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    foreach (DataRow item in statListTable.Rows)
                    {
                        int id = item.Field<int>(statListTable.Columns["Id"]);
                        string filename = statlist.Substring(0, statlist.LastIndexOf('\\') + 1) + id.ToString() + ".txt";
                        try
                        {
                            using (StreamReader sr = new StreamReader(filename, Encoding.GetEncoding(1251)))
                            {
                                string line;
                                while ((line = sr.ReadLine()) != null)
                                {
                                    ParseWrTable(line);
                                }
                            }
                        }
                        catch (IOException ex)
                        {
                            MessageBox.Show("Файл не мог быть прочитан:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                
                SendMessage(dataGridView1.Handle, WM_SETREDRAW, false, 0);
                dataGridView1.DataSource = wrTable;
                dataGridView1.Columns["Id"].Visible = false;
                dataGridView1.Columns["StatId"].HeaderText = "Индекс ВМО";
                dataGridView1.Columns["Name"].HeaderText = "Станция";
                dataGridView1.Columns["Year"].HeaderText = "Год";
                dataGridView1.Columns["Jan"].Visible = false;
                dataGridView1.Columns["Feb"].Visible = false;
                dataGridView1.Columns["Mar"].Visible = false;
                dataGridView1.Columns["Apr"].Visible = false;
                dataGridView1.Columns["May"].Visible = false;
                dataGridView1.Columns["Jun"].Visible = false;
                dataGridView1.Columns["Jul"].Visible = false;
                dataGridView1.Columns["Aug"].Visible = false;
                dataGridView1.Columns["Sep"].Visible = false;
                dataGridView1.Columns["Oct"].Visible = false;
                dataGridView1.Columns["Nov"].Visible = false;
                dataGridView1.Columns["Dec"].Visible = false;
                foreach (string month in GetMonthList())
                {
                    dataGridView1.Columns[month].HeaderText = MonthToString(month);
                    dataGridView1.Columns[month].Visible = true;
                }
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                SendMessage(dataGridView1.Handle, WM_SETREDRAW, true, 0);
                dataGridView1.Refresh();
                button2.Enabled = true;
            }
            else
            {
                MessageBox.Show("Не обнаружены нужные файлы в выбранной папке", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<int> GetStatList()
        {
            List<int> list = new List<int>();
            foreach (DataRow item in statListTable.Rows)
            {
                list.Add(item.Field<int>(statListTable.Columns["Id"]));
            }
            return list;
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

        private string StringToMonth(string str)
        {
            switch (str)
            {
                case "Январь": return "Jan";
                case "Февраль": return "Feb";
                case "Март": return "Mar";
                case "Апрель": return "Apr";
                case "Май": return "May";
                case "Июнь": return "Jun";
                case "Июль": return "Jul";
                case "Август": return "Aug";
                case "Сентябрь": return "Sep";
                case "Октябрь": return "Oct";
                case "Ноябрь": return "Nov";
                case "Декабрь": return "Dec";
                default: return "";
            }
        }

        private List<string> GetMonthList()
        {
            List<string> list = new List<string>();
            foreach (DataRow item in fldTable.Rows)
            {
                string col_name = StringToMonth(item.Field<string>(fldTable.Columns["ColName"]));
                if (!String.IsNullOrEmpty(col_name))
                {
                    list.Add(col_name);
                }
            }
            return list;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2(wrTable, GetStatList(), GetMonthList());
            form2.ShowDialog();
        }
    }
}
