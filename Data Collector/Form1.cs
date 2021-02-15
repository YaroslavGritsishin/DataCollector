using Data_Collector.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
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
            int startCycle = Convert.ToInt32(comboBox1.Text);
            int endCycle = Convert.ToInt32(comboBox2.Text);

            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                string res = string.Empty;
                Groups.TryGetValue(checkedListBox1.CheckedItems[i].ToString(), out res);

                for (; startCycle < endCycle + 1; )
                {
                    if(comboBox3.Text == "План")
                    {
                        if(comboBox4.Text == "Горизонтальное")
                        {
                            var pointData = WorkWithDB.GetHorizontalPositionPoints(startCycle, WorkWithDB.GetAllPointsName(res));
                            var table = Word.CreteHorizontalPositionTable(pointData, Convert.ToInt32(comboBox1.Text), 0.4);
                            Word.AddTableInBookMark("C:\\1\\12_Horizontal.docx", table, "Вставка");
                            startCycle += 3;
                        }
                        if (comboBox4.Text == "Вертикальное")
                        {
                            var pointData = WorkWithDB.GetVerticalPositionPoints(startCycle, WorkWithDB.GetAllPointsName(res));
                            var table = Word.CreateVerticalPositionTable (pointData);
                            Word.AddTableInBookMark("C:\\1\\12.docx", table, "Вставка");
                            startCycle++;
                        }
                    }
                    if(comboBox3.Text == "Высота")
                    {
                        if(comboBox4.Text == "Горизонтальное")
                        {

                        }
                        if (comboBox4.Text == "Вертикальное")
                        {

                        }

                    }
                    
                }
                
            }
        }
    }
}
