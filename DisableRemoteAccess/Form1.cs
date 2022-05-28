
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
            openXlsx();
        }

        public void openXlsx()
        {
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
                label1.Text = "Не удалось открыть файл! Возмжные причины:\n 1. Файл перемещен. \n 2. Не заполнена одна из требуемых строк для проверки оплаты.";
            }

         //   Thread.Sleep(30000);
         //   openXlsx();
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


        private void label1_Click(object sender, EventArgs e)
        {
            /*using (ITikConnection connection = ConnectionFactory.CreateConnection(TikConnectionType.Api))
            {
             //   connection.Open(HOST, USER, PASS);

                int n = 1;

                var natRule = connection.CreateCommandAndParameters("/ip/route/print", "dst-address", Action.route).ExecuteList();
                var value = natRule.Count();
                if (value == n)
                {
                    Color colorOn = Color.Blue;
                    label1.ForeColor = colorOn; 
                    label1.Text = ("Текущий статус сервера\n          ДОСТУПЕН!");
                }
            }
            using (ITikConnection connection = ConnectionFactory.CreateConnection(TikConnectionType.Api))
            {              
            //    connection.Open(HOST, USER, PASS);
              
                int i = 0;
                
                var natRule = connection.CreateCommandAndParameters("/ip/route/print", "dst-address", Action.route).ExecuteList();
                var value = natRule.Count();
                if (value == i)
                {
                    Color colorOff_n = Color.Red;
                    label1.ForeColor = colorOff_n;
                    label1.Text = ("Текущий статус сервера\n        НЕ ДОСТУПЕН!");
                }
            }*/
        }
    }
}
