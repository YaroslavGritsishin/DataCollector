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
        public Form1()
        {
            InitializeComponent();
        }

        private void grab_btn_Click(object sender, EventArgs e)
        {
            ExApp = new Excell.Application() {DisplayAlerts = false };
            Excel Ex = new Excel(ExApp);
            Ex.GetData();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            ExApp?.Quit();
            foreach(var process in Process.GetProcesses())
            {
                if (process.ProcessName == "EXCEL") process.Kill();
            }
        }
    }
}
