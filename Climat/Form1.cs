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
using System.Text.RegularExpressions;

namespace Climat
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, Int32 wMsg, bool wParam, Int32 lParam);
        private const int WM_SETREDRAW = 11;
        public static DataSet climatSet;
        public static DataTable statListTable, fldTable, wrTable;
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
                DataColumn IdColumn = new DataColumn("ID", Type.GetType("System.Int32"));
                IdColumn.Unique = true;
                IdColumn.AllowDBNull = false;

                statListTable.Columns.Add(IdColumn);
                statListTable.Columns.Add(new DataColumn("Название станции", Type.GetType("System.String")));
                statListTable.PrimaryKey = new DataColumn[] { statListTable.Columns["ID"] };
            }

            {
                DataColumn IdColumn = new DataColumn("ID", Type.GetType("System.Int32"));
                IdColumn.Unique = true;
                IdColumn.AllowDBNull = false; 

                fldTable.Columns.Add(IdColumn);
                fldTable.Columns.Add(new DataColumn("N", Type.GetType("System.Int32")));
                fldTable.Columns.Add(new DataColumn("Формат", Type.GetType("System.String")));
                fldTable.Columns.Add(new DataColumn("Название столбца", Type.GetType("System.String")));
                fldTable.PrimaryKey = new DataColumn[] { fldTable.Columns["ID"] };
            }

            {
                DataColumn IdColumn = new DataColumn("ID", Type.GetType("System.Int32"));
                IdColumn.Unique = true;
                IdColumn.AllowDBNull = false;
                IdColumn.AutoIncrement = true;
                IdColumn.AutoIncrementSeed = 1;
                IdColumn.AutoIncrementStep = 1;

                wrTable.Columns.Add(IdColumn);
                wrTable.Columns.Add(new DataColumn("Индекс ВМО", Type.GetType("System.Int32")));
                wrTable.Columns.Add(new DataColumn("Название станции", Type.GetType("System.String")));
                wrTable.PrimaryKey = new DataColumn[] { wrTable.Columns["ID"] };
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

        public static string final_path = "";

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
                while (wrTable.Columns.Count > 4)
                {
                    wrTable.Columns.RemoveAt(wrTable.Columns.Count - 1);
                }
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
                            if (id == 1)
                            {
                                wrTable.Columns["Индекс ВМО"].ColumnName = colName;
                            }
                            else { 
                                wrTable.Columns.Add(new DataColumn(colName, Type.GetType("System.String")));
                            }
                        }
                    }
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Файл не мог быть прочитан:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                if (wr.Length > 0)
                {
                    ParseWrTable(wr);
                }
                else
                {
                    foreach (DataRow item in statListTable.Rows)
                    {
                        int id = item.Field<int>(statListTable.Columns["ID"]);
                        string filename = statlist.Substring(0, statlist.LastIndexOf('\\') + 1) + id.ToString() + ".txt";
                        ParseWrTable(filename);
                    }
                }
                
                SendMessage(dataGridView1.Handle, WM_SETREDRAW, false, 0);
                dataGridView1.DataSource = wrTable;
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.AutoResizeColumnHeadersHeight();
                dataGridView1.AutoResizeColumns();
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.Columns["Название станции"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                SendMessage(dataGridView1.Handle, WM_SETREDRAW, true, 0);
                dataGridView1.Refresh();
                button2.Enabled = true;
                Form1.final_path = path;
            }
            else
            {
                MessageBox.Show("Не обнаружены нужные файлы в выбранной папке", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ParseWrTable(string path)
        {
            try
            {
                using (StreamReader sr = new StreamReader(path, Encoding.GetEncoding(1251)))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        DataRow row = wrTable.NewRow();
                        int startIndex = 0;
                        for (int i = 1; i <= fldTable.Rows.Count; i++)
                        {
                            string format = fldTable.Select("ID = " + i.ToString()).First()["Формат"].ToString();
                            string field_name = fldTable.Select("ID = " + i.ToString()).First()["Название столбца"].ToString();
                            int length = Convert.ToInt32(Regex.Replace(format, @"\,\d+", ""));
                            row[field_name] = line.Substring(startIndex, length).Trim();
                            startIndex += length + 1;
                        }
                        row[2] = statListTable.Select("ID = " + row[1].ToString()).First()["Название станции"].ToString();
                        wrTable.Rows.Add(row);
                    }
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show("Файл не мог быть прочитан:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }
    }
}