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
            var originalPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\original.xlsx";
            var testFilePath = System.Environment.GetFolderPath(
                Environment.SpecialFolder.DesktopDirectory) + 
                @"\updated_" + 
                DateTime.Now.ToString("gyyyy年MM月dd日ddddtthh時mm分ss") + 
                ".xlsx";
            System.IO.File.Copy(
                originalPath,
                testFilePath,
                true
                );
            var a = new ClosedXML.Excel.XLWorkbook(testFilePath);
            var ws = a.Worksheet(1);
            ws.Cell("A4").Value = textBox1.Text;
            a.Save();
            MessageBox.Show(testFilePath + "を保存しました。");
        }
    }
}
