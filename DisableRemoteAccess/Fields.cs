using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DisableRemoteAccess
{
    public static class Fields
    {
        public static string fileXlsx = @"C:\Data\Clients.xlsx";     
        public static string customerName;                           
        public static string paymentState;
        public static string serverName;
        public static string addrNameRule;

        public static int customerNameRow = 1;                       
        public static int paymentStateRow = 1;                                                          
        public static int serverNameRow = 1;                          
        public static int addrNameRuleRow = 1;

        public static string addrListON;
        public static string addrListOFF;
    }
}
