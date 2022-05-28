using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using tik4net;

namespace DisableRemoteAccess
{
    public static class Action
    {
        public static string USER = "apiUser";
        public static string PASS = "Password";

        public static void enable()
        {
                using (ITikConnection connection = ConnectionFactory.CreateConnection(TikConnectionType.Api))
                {
                    connection.Open(Fields.serverName, USER, PASS);

                    var natRule = connection.CreateCommandAndParameters("/ip/address/print", "comment", Fields.addrNameRule).ExecuteList();
                    var id = natRule.Single().GetId();
                    var disableRule = connection.CreateCommandAndParameters("/ip/address/enable", TikSpecialProperties.Id, id);
                    disableRule.ExecuteNonQuery();
                }
        }

        public static void disable()
        {
            using (ITikConnection connection = ConnectionFactory.CreateConnection(TikConnectionType.Api))
            {
                connection.Open(Fields.serverName, USER, PASS);

                var natRule = connection.CreateCommandAndParameters("/ip/address/print", "comment", Fields.addrNameRule).ExecuteList();
                var id = natRule.Single().GetId();
                var enableRule = connection.CreateCommandAndParameters("/ip/address/disable", TikSpecialProperties.Id, id);
                enableRule.ExecuteNonQuery();
            }
        }
    }
}
