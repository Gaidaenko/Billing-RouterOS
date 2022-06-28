
using System;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Windows.Forms;
using Color = System.Drawing.Color;
using System.Diagnostics;
using System.Threading.Tasks;
using System.IO;

namespace DisableRemoteAccess
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Run();
        }
        public async Task Run()
        {
            while (true)
            {
                openXlsx();
                closeXlsx();
                await Task.Delay(TimeSpan.FromSeconds(15));
            }
        }
        public async Task openXlsx()
        {
            Color launched = Color.Green;
            label1.ForeColor = launched;
            label1.Text = "Мониторинг оплаты запущен!";
            label3.Text = null + "Должники:\n";
            label4.Text = null;

            if (File.Exists(Fields.fileXlsx))
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

                Range RngClients, RngState, RngServer, RngRule, RngMail;

                RngClients = xlsApp.get_Range("A2", "A100");                           
                var dataArrClients = (object[,])RngClients.Value;

                RngState = xlsApp.get_Range("B2", "B100");                           
                var dataArrState = (object[,])RngState.Value;

                RngServer = xlsApp.get_Range("C2", "C100");                           
                var dataArrServer = (object[,])RngServer.Value;

                RngRule = xlsApp.get_Range("D2", "D100");                            
                var dataArrRule = (object[,])RngRule.Value;

                RngMail = xlsApp.get_Range("E2", "E100");                                              
                var dataArrMail = (object[,])RngMail.Value;                                 

                while (dataArrClients[Fields.customerNameRow, 1] != null && dataArrState[Fields.addrNameRuleRow, 1] != null && 
                       dataArrServer[Fields.addrNameRuleRow, 1] != null && dataArrMail[Fields.addrMailRow, 1] != null && dataArrMail[Fields.addrMailRow, 1] != null)
                {
                    Fields.customerName = dataArrClients[Fields.customerNameRow, 1].ToString();
                    Fields.paymentState = dataArrState[Fields.paymentStateRow, 1].ToString();
                    Fields.serverName = dataArrServer[Fields.serverNameRow, 1].ToString();
                    Fields.addrNameRule = dataArrRule[Fields.addrNameRuleRow, 1].ToString();
                    Fields.addrMail = dataArrMail[Fields.addrMailRow, 1].ToString();

                    if (Fields.paymentState == "Оплачено")
                    {
                        Action.enable();
                    }
                    else
                    {
                        label3.Text += "\n " + Fields.customerName;
                        Action.disable();
                    }

                    Fields.customerNameRow++;
                    Fields.paymentStateRow++;
                    Fields.addrNameRuleRow++;
                    Fields.serverNameRow++;
                    Fields.addrMailRow++;
                }

                if (Fields.сonnectionError != 0)
                {
                    Color connError = Color.Red;
                    label4.ForeColor = connError;
                    label4.Text = "Присутствуют шлюзы к которым нельзя подключится.\nДля большей информации смотрите лог Windows, Billing!";
                }

                ObjWorkBook.Close();
                xlsApp.Application.Quit();
                xlsApp = null;
                ObjWorkBook = null;
                dataArrClients = null;
                dataArrState = null;
                dataArrServer = null;
                dataArrRule = null;
                dataArrMail = null;
            }
            if (!File.Exists(Fields.fileXlsx))
            {
                Color warning = Color.Red;
                label1.ForeColor = warning;
                label1.Text = "Не удалось запустить мониторинг!\n1. Возможно фай xlsx перемещен или переименован. \n2. Не заполнена одна из требуемых строк для проверки оплаты.";              
                return;
            }
        }
        public void closeXlsx()
        {
              Process List = Process.GetProcessesByName("EXCEL").Last();         
              Fields.customerNameRow = 1;
              Fields.paymentStateRow = 1;
              Fields.addrNameRuleRow = 1;
              Fields.serverNameRow = 1;
              Fields.addrMailRow = 1;
              Fields.сonnectionError = 0;
        }
        private void button1_Click(object sender, EventArgs e)
        {
              Fields.customerNameRow = 1;
              Fields.paymentStateRow = 1;
              Fields.addrNameRuleRow = 1;
              Fields.serverNameRow = 1;
              Fields.addrMailRow = 1;
              Fields.сonnectionError = 0;
              openXlsx();
        }
        private void adoutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Billing VMs v1.0");
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void label2_Click(object sender, EventArgs e)
        {
            
        }
        private void label3_Click(object sender, EventArgs e)
        {
            
        }
        public void label4_Click(object sender, EventArgs e)
        {

        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
