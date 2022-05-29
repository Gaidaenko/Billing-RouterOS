
using System;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Windows.Forms;
using tik4net;
using Color = System.Drawing.Color;
using System.Diagnostics;
using System.Threading;

namespace DisableRemoteAccess
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
           // openXlsx();
        }


        public void openXlsx()
        {
            Color launched = Color.Green;
            label1.ForeColor = launched;
            label1.Text = "Мониторинг оплаты запущен.";

            try
            {
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
            Next:
                Range Rng, CheckingRow;
                Rng = xlsApp.get_Range("A2", "F34");
                var dataArr = (object[,])Rng.Value;

                if (dataArr[Fields.customerNameRow, Fields.paymentStateRow] != null && dataArr[Fields.addrNameRuleRow, Fields.serverNameRow] != null)
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

                    goto Next;
                }
                else
                {
                    killProcess();
                }
            }
            catch (Exception s)
            {
                Color warning = Color.Red;
                label1.ForeColor = warning;
                label1.Text = "Не удалось открыть файл! Возмжные причины:\n1. Файл xlsx перемещен или переименован. \n2. Не заполнена одна из требуемых строк для проверки оплаты.";
            }


            
        }
        public void killProcess()
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (var process in List)
            {
                process.Kill();
            }
        }
        public void label1_Click(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            openXlsx();
     
        }
    }
}
