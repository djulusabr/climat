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

namespace Climat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Multiselect = false;
            dlg.Filter = "Archives (*.rar;*.zip;*.7z)|*.rar;*.zip;*.7z";
            
            if (dlg.ShowDialog() == DialogResult.OK) 
            {
                label1.Text = dlg.FileName;
                button2.Enabled = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var dlg = new FolderSelectDialog();

            if (dlg.ShowDialog(IntPtr.Zero))
            {
                var fileList = Directory.GetFiles(dlg.FileName, "*.txt");
                string fld = "", statlist = "", wr = "";
                foreach (string str in fileList) {
                    if (str.Contains("fld"))
                        fld = str;
                    else if (str.Contains("statlist"))
                        statlist = str;
                    else if (str.Contains("wr"))
                        wr = str;
                }
                if (fld.Length > 0 && statlist.Length > 0 && wr.Length > 0)
                {
                    try
                    {
                        using (StreamReader sr = new StreamReader(fld))
                        {
                            string line = sr.ReadLine();
                        }
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show("Файл не мог быть прочитан:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    button2.Enabled = true;
                }
                else
                {
                    MessageBox.Show("Не обнаружены нужные файлы в выбранной папке", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
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
