﻿
using System;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Windows.Forms;
using tik4net;
using Color = System.Drawing.Color;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Timer = System.Windows.Forms.Timer;

namespace DisableRemoteAccess
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            openXlsx();
        }
        public async Task openXlsx()
        {
            Color launched = Color.Green;
            label1.ForeColor = launched;
            label1.Text = "Мониторинг оплаты запущен!";

            try
            {
                DateTime dateTime = DateTime.Now;
                label2.Text = "Время запуска проверки\n " + dateTime.ToString();

                Excel.Application xlsApp = new Excel.Application();               
                Workbook ObjWorkBook = xlsApp.Workbooks.Open
                                    (Filename: Fields.fileXlsx,
                                     UpdateLinks: 0,
                                     ReadOnly: true,
                                     Format: 5,
                                     false,
                                     Origin: XlPlatform.xlWindows,
                                     Delimiter: "",
                                     Editable: true,
                                     Notify: false,
                                     Converter: 0,
                                     AddToMru: true,
                                     Local: false,
                                     CorruptLoad: false);

                Worksheet worksheet = xlsApp.Worksheets[1];
                worksheet.Activate();

                Range Rng, CheckingRow;
                Rng = xlsApp.get_Range("A2", "F34");
                var dataArr = (object[,])Rng.Value;

                while (dataArr[Fields.customerNameRow, Fields.paymentStateRow] != null && dataArr[Fields.addrNameRuleRow, Fields.serverNameRow] != null)
                {
                    Fields.customerName = dataArr[Fields.customerNameRow, 3].ToString();
                    Fields.paymentState = dataArr[Fields.paymentStateRow, 4].ToString();

                    Fields.serverName = dataArr[Fields.serverNameRow, 5].ToString();
                    Fields.addrNameRule = dataArr[Fields.addrNameRuleRow, 6].ToString();

                    if (Fields.paymentState == "Оплачено")
                    {
                        Action.enable();
                    }
                    else
                    {
                        Action.disable();
                    } 

                    Fields.customerNameRow++;
                    Fields.paymentStateRow++;
                    Fields.addrNameRuleRow++;
                    Fields.serverNameRow++;
                }
           
                ObjWorkBook.Close();
                xlsApp.Application.Quit();
                xlsApp = null;
                ObjWorkBook = null;
                dataArr = null;
                Rng = null;               

                closeXlsx();
            }
            catch (Exception s)
            {
                Color warning = Color.Red;
                label1.ForeColor = warning;
                label1.Text ="Не удалось запустить мониторинг!\n1. Возможно фай xlsx открыт, перемещен или переименован. \n2. Не заполнена одна из требуемых строк для проверки оплаты.";
            }

            await Task.Delay(TimeSpan.FromMinutes(5));
            openXlsx();
        }

        public void closeXlsx()
        {
              Process List = Process.GetProcessesByName("EXCEL").Last();
           
              Fields.customerNameRow = 1;
              Fields.paymentStateRow = 1;
              Fields.addrNameRuleRow = 1;
              Fields.serverNameRow = 1;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            openXlsx();
        }
        private void adoutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Billing VMs v1.0");
        }
        public void label1_Click(object sender, EventArgs e)
        {

        }
        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
