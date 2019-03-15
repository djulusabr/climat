using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Compression;

namespace Climat
{
    public partial class Form1 : Form
    {
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

                DataColumn StatIdColumn = new DataColumn("StatId", Type.GetType("System.Int32"));
            
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
                Console.WriteLine(fileName);
            }
            if (statlist.Length > 0 && fld.Length > 0)
            {
                try
                {
                    using (StreamReader sr = new StreamReader(statlist, Encoding.GetEncoding(1251)))
                    {
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            int id = Convert.ToInt32(line.Substring(0, 5));
                            string name = line.Substring(6, line.Length - 6);
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
                            string colName = line.Substring(15, line.Length - 15);
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
                                int id = Convert.ToInt32(line.Substring(0, 5));
                                int year = Convert.ToInt32(line.Substring(6, 4));
                                string jan = line.Substring(11, 5).Replace(" ", string.Empty);
                                string feb = line.Substring(17, 5).Replace(" ", string.Empty);
                                string mar = line.Substring(23, 5).Replace(" ", string.Empty);
                                string apr = line.Substring(29, 5).Replace(" ", string.Empty);
                                string may = line.Substring(35, 5).Replace(" ", string.Empty);
                                string jun = line.Substring(41, 5).Replace(" ", string.Empty);
                                string jul = line.Substring(47, 5).Replace(" ", string.Empty);
                                string aug = line.Substring(53, 5).Replace(" ", string.Empty);
                                string sep = line.Substring(59, 5).Replace(" ", string.Empty);
                                string oct = line.Substring(65, 5).Replace(" ", string.Empty);
                                string nov = line.Substring(71, 5).Replace(" ", string.Empty);
                                string dec = line.Substring(77, 5).Replace(" ", string.Empty);
                                wrTable.Rows.Add(new object[] { id, statid, year, jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec });
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

                }

                dataGridView1.DataSource = wrTable;
                MessageBox.Show("Прочитано успешно", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button2.Enabled = true;
            }
            else
            {
                MessageBox.Show("Не обнаружены нужные файлы в выбранной папке", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
            string ctnName = "richTextBox1";
            Control ctn = form2.Controls[ctnName];
            ctn.Text = label1.Text;
        }
    }
}
