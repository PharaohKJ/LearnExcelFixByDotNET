using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var testFilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\test.xlsx";
            System.IO.File.Copy(
                @"C:\Users\真透\Downloads\領収書テンプレート.xlsx",
                testFilePath,
                true
                );
            var a = new ClosedXML.Excel.XLWorkbook(testFilePath);
            var ws = a.Worksheet(1);
            MessageBox.Show(ws.Cell("A4").Value.ToString());
            ws.Cell("A4").Value = "ファランクスウェア";
            a.Save();
        }
    }
}
