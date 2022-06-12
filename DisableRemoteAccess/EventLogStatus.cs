using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DisableRemoteAccess
{
    public static class EventLogStatus
    {
        public static void eventLogEnabled()
        {           
            string Event = "Внимание: доступ к серверу "+ Fields.customerName + " отключен.";
            string Source = "Access enabled";
            string Log = "Billing";

            if (!EventLog.SourceExists(Source))
                EventLog.CreateEventSource(Source, Log);
            using (EventLog eventLog = new EventLog("Billing"))
            {
                eventLog.Source = "Access disabled";                       
                eventLog.WriteEntry(Event, EventLogEntryType.Warning);             
                return;
            }
        }

        public static void eventLogDisabled()
        { 
            string Event = "Внимание: доступ к серверу "+ Fields.customerName + " включен.";
            string Source = "Access enabled";
            string Log = "Billing";

            if (!EventLog.SourceExists(Source))
                EventLog.CreateEventSource(Source, Log);
            using (EventLog eventLog = new EventLog("Billing"))
            {
                eventLog.Source = "Access enabled";                        
                eventLog.WriteEntry(Event, EventLogEntryType.Information);
                return;
            }
        }

        public static void noAccessToGateway()
        {
            string Event = "Не удалось подключится к " + Fields.serverName;
            string Source = "Тo access to gateway";
            string Log = "Billing";

            if (!EventLog.SourceExists(Source))
                EventLog.CreateEventSource(Source, Log);
            using (EventLog eventLog = new EventLog("Billing"))
            {
                eventLog.Source = "Тo access to gateway";                       
                eventLog.WriteEntry(Event, EventLogEntryType.Error);
                return;
            }
        }
    }
}
