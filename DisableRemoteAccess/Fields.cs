using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DisableRemoteAccess
{
    public static class Fields
    {

        public static string fileXlsx = @"C:\Data\Clients.xlsx";     //Пусть к файлу xlsx
        public static string customerName;                           // Колонка - имя клиента 
        public static string paymentState;                           // Колонка - статус оплаты
        public static int customerNameRow = 1;                       // Отсчет итерации, колонка - клиенты
        public static int paymentStateRow = 1;                       // Отсчет итерации, колонка - статус оплаты

        public static string serverName;                             // Имя сервера
        public static string addrNameRule;                               // Имя Address List
        public static int serverNameRow =1;                          // Отсчет итерации, колонка - имя сервера
        public static int addrNameRuleRow =1;                            // Отсчет итерации, колонка - имя Address List на шлюзе

    }
}
