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
        public static string USER = "login";
        public static string PASS = "apiPassword";       
        public static void enable()
        {
            try
            {
                using (ITikConnection connection = ConnectionFactory.CreateConnection(TikConnectionType.Api))
                {
                    connection.Open(Fields.serverName, USER, PASS);

                    var natRule = connection.CreateCommandAndParameters("/ip/address/print", "comment", Fields.addrNameRule).ExecuteList();
                    var id = natRule.Single().GetId();
                    var disableRule = connection.CreateCommandAndParameters("/ip/address/enable", TikSpecialProperties.Id, id);
                    disableRule.ExecuteNonQuery();

                    foreach (var result in natRule)
                    {
                        var items = result.Words.GetEnumerator();

                        while (items.MoveNext())
                        {
                                
                            if (items.Current.Key == "disabled" && items.Current.Value == "true")               
                            {
                                // shutdown notice
                            }
                        }
                    } 
                }
            }
            catch (Exception e)
            {

                //  connection exception
            }
        }

        public static void disable()
        {
            try
            {
                using (ITikConnection connection = ConnectionFactory.CreateConnection(TikConnectionType.Api))
                {
                    connection.Open(Fields.serverName, USER, PASS);

                    var natRule = connection.CreateCommandAndParameters("/ip/address/print", "comment", Fields.addrNameRule).ExecuteList();
                    var id = natRule.Single().GetId();
                    var enableRule = connection.CreateCommandAndParameters("/ip/address/disable", TikSpecialProperties.Id, id);
                    enableRule.ExecuteNonQuery();

                    foreach (var result in natRule)
                    {
                        var items = result.Words.GetEnumerator();

                        while (items.MoveNext())
                        {
                            if (items.Current.Key == "disabled" && items.Current.Value == "false")             
                            {
                                //activation notice
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {

                // connection exception

            }
        }
    }
}
