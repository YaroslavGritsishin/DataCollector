using Data_Collector.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excell = Microsoft.Office.Interop.Excel;

namespace Data_Collector
{
    public partial class Form1 : Form
    {
        Excell.Application ExApp;
        Dictionary<string, string> Groups;
        public Form1()
        {
            InitializeComponent();
            Groups = new Dictionary<string, string>();
            Groups.Add(checkedListBox1.Items[0].ToString(), "BP");
            Groups.Add(checkedListBox1.Items[1].ToString(), "ZP");
            Groups.Add(checkedListBox1.Items[2].ToString(), "P");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            ExApp?.Quit();
            foreach (var process in Process.GetProcesses())
            {
                if (process.ProcessName == "EXCEL") process.Kill();
            }
        }

        private void собратьДанныеСExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            ExApp = new Excell.Application() { DisplayAlerts = false };
            Excel Ex = new Excel(ExApp);
            Ex.GetData();
            button1.Enabled = true;
        }
        private void ComboboxFiling(ComboBox comboBox)
        {
            using (Context context = new Context())
            {
                int position = 0;
                object[] items = new object[context.pointCoordinates.Select(point => point.CycleNumber).GroupBy(x => x).Count()];
                foreach (var item in context.pointCoordinates.Select(point => point.CycleNumber).GroupBy(x => x).ToList())
                {
                    items[position] = item.Key;
                    position++;
                }
                comboBox.Items.AddRange(items);
                comboBox.SelectedIndex = 0;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ComboboxFiling(comboBox1);
            ComboboxFiling(comboBox2);
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            string HorizontalTemplatePath;
            string VerticalTemplatePath;
            int _startCycle = Convert.ToInt32(comboBox1.Text);
            int _endCycle = Convert.ToInt32(comboBox2.Text);
            int startCycle = 0;
            int endCycle = 0;
            int FirstMachCycle = 0;
            if (_startCycle > _endCycle)
            {
                if (MessageBox.Show("Не верно указан диапазон", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    comboBox2.Text = comboBox1.Text;
                    return;
                }
            }
            if (!Directory.Exists("C:\\1\\Сводные ведомости"))
                Directory.CreateDirectory("C:\\1\\Сводные ведомости");
            if (!File.Exists("C:\\1\\VerticalTemplate.docx"))
                if (MessageBox.Show(@"Поместите шаблон файла Word в папку C:\1 c названием VerticalTemplate.docx", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.OK) return;
            if (!File.Exists("C:\\1\\HorizontalTemplate.docx"))
                if (MessageBox.Show(@"Поместите шаблон файла Word в папку C:\1 c названием HorizontalTemplate.docx", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.OK) return;

            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                startCycle = _startCycle;
                endCycle = _endCycle;
                string res = string.Empty;
                Groups.TryGetValue(checkedListBox1.CheckedItems[i].ToString(), out res);
                using (Context context = new Context())
                {
                    FirstMachCycle = context.pointCoordinates.Where(point => point.Name.StartsWith(res)).OrderBy(point => point.CycleNumber).First().CycleNumber;
                }
                if (FirstMachCycle > endCycle) continue;
                CopyTemplateFile(comboBox3, comboBox4, out HorizontalTemplatePath, out VerticalTemplatePath, checkedListBox1.CheckedItems[i].ToString());
                for (; startCycle < endCycle + 1;)
                {
                    if (FirstMachCycle > startCycle)
                    {
                        startCycle++;
                        continue;
                    }
                    else FirstMachCycle = 0;
                    
                    if (comboBox3.Text == "План")
                    {
                        if (comboBox4.Text == "Горизонтальное")
                        {
                            var pointData = WorkWithDB.GetHorizontalPositionPoints(startCycle, WorkWithDB.GetAllPointsName(res));
                            if (pointData.Count != 0)
                            {
                                var table = Word.CreteHorizontalPositionTable(pointData, startCycle, 0.4);
                                Word.AddTableInBookMark(HorizontalTemplatePath, table, "Вставка");
                            }
                            startCycle += 3;
                        }
                        if (comboBox4.Text == "Вертикальное")
                        {
                            var pointData = WorkWithDB.GetVerticalPositionPoints(startCycle, WorkWithDB.GetAllPointsName(res));
                            if (pointData.Count != 0)
                            {
                                var table = Word.CreateVerticalPositionTable(pointData);
                                Word.AddTableInBookMark(VerticalTemplatePath, table, "Вставка");
                            }
                            startCycle++;
                        }
                    }
                    if (comboBox3.Text == "Высота")
                    {
                        if (comboBox4.Text == "Горизонтальное")
                        {

                            var pointData = WorkWithDB.GetHorizontalElivationPoints(startCycle, WorkWithDB.GetAllPointsName(res));
                            if (pointData.Count != 0)
                            {
                                var table = Word.CreteHorizontalElevationTable(pointData, startCycle, 0.4);
                                Word.AddTableInBookMark(HorizontalTemplatePath, table, "Вставка");
                            }
                            startCycle += 6;
                        }
                        if (comboBox4.Text == "Вертикальное")
                        {
                            var pointData = WorkWithDB.GetVerticalElivationPoints(startCycle, WorkWithDB.GetAllPointsName(res));
                            if (pointData.Count != 0)
                            {
                                var table = Word.CreateVerticalElivationTable(pointData, startCycle);
                                Word.AddTableInBookMark(VerticalTemplatePath, table, "Вставка");
                            }

                            startCycle += 3;
                        }

                    }

                }

            }
            button1.Enabled = true;
        }

        private void CopyTemplateFile(ComboBox typeCoordinate, ComboBox positionOnSheets, out string HorizontalTemplatePath, out string VerticalTemplatePath, string pointTypeName)
        {
            HorizontalTemplatePath = string.Empty;
            VerticalTemplatePath = string.Empty;

            if (typeCoordinate.Text == "План")
            {
                if (positionOnSheets.Text == "Горизонтальное")
                {
                    HorizontalTemplatePath = $"C:\\1\\Сводные ведомости\\Горизонтальная сводная ведомость планового положения({pointTypeName})_{ConvertToMD5(Guid.NewGuid().ToString())}.docx";
                    File.Copy("C:\\1\\HorizontalTemplate.docx", HorizontalTemplatePath);
                }
                if (positionOnSheets.Text == "Вертикальное")
                {

                    VerticalTemplatePath = $"C:\\1\\Сводные ведомости\\Вертикальна сводная ведомость планового положения({pointTypeName})_{ConvertToMD5(Guid.NewGuid().ToString())}.docx";
                    File.Copy("C:\\1\\VerticalTemplate.docx", VerticalTemplatePath);
                }
            }
            if (typeCoordinate.Text == "Высота")
            {
                if (positionOnSheets.Text == "Горизонтальное")
                {
                    HorizontalTemplatePath = $"C:\\1\\Сводные ведомости\\Горизонтальная сводная ведомость высотного положения({pointTypeName})_{ConvertToMD5(Guid.NewGuid().ToString())}.docx";
                    File.Copy("C:\\1\\HorizontalTemplate.docx", HorizontalTemplatePath);
                }
                if (positionOnSheets.Text == "Вертикальное")
                {

                    VerticalTemplatePath = $"C:\\1\\Сводные ведомости\\Вертикальна сводная ведомость высотного положения({pointTypeName})_{ConvertToMD5(Guid.NewGuid().ToString())}.docx";
                    File.Copy("C:\\1\\VerticalTemplate.docx", VerticalTemplatePath);
                }
            }
        }
        private string ConvertToMD5(string value)
        {
            MD5 md5 = MD5.Create();
            byte[] inputBytes = Encoding.ASCII.GetBytes(value);
            byte[] hash = md5.ComputeHash(inputBytes);
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < hash.Length; i++)
            {
                builder.Append(hash[i].ToString("x2"));
            }
            return builder.ToString();
        }
    }
}
