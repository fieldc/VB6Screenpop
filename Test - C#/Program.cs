using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Remoting;
using System.ComponentModel;
using System.Threading;
using System.Diagnostics;
using SSFIntegration;

namespace CSComTest
{
    [ComImport(),Guid(SSFIntegration.SSFComMessageProxy.ClassId)]
    class _SSFIntegrationCOMObject
    {
    }

    public class EventSink : SSFIntegration.IRecordRequestCallback
    {
        void IRecordRequestCallback.Callback(string clientId, string taxId, string contractId, int contractCount)
        {
            System.Console.WriteLine("\t[IRecordRequestCallback.Callback] got event: clientId=" + clientId + " taxId="+taxId+" contractId=" + contractId+" contractCount="+contractCount);
        }
    }

    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            _SSFIntegrationCOMObject com_obj = new _SSFIntegrationCOMObject();
            SSFComMessageProxy co = new SSFComMessageProxy();
            IRecordRequest obj = (IRecordRequest)com_obj;
            co.Requested += new SSFComMessageProxy.RequestedEventHandler(co_Requested);
            /*
            EventSink es = new EventSink();
            try
            {
                obj.RegisterCallBack(es);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Debug.WriteLine(ex.StackTrace);
            }*/
            /*
            System.Console.WriteLine("CSComTest.Pid: " + Process.GetCurrentProcess().Id);
            uint a = 0, b = 0;
            obj.GetProcessThreadID(out a, out b);
            System.Console.WriteLine("obj.processId=" + a + " threadId=" + b);
            */
            bool docontinue=true;
            string clientId = "";
            string taxId = "";
            string contractId = "";
            string numPolicies = "";
            do{

                System.Console.WriteLine("Press 'q' to quit");
                System.Console.Write("ClientId: ");
                clientId = System.Console.ReadLine().Trim();
                if (clientId.ToLower() == "q")
                {
                    docontinue = false;
                    break;
                }

                System.Console.Write("TaxId: ");
                taxId = System.Console.ReadLine().Trim();
                if (taxId.ToLower() == "q")
                {
                    docontinue = false;
                    break;
                }
                System.Console.Write("ContractId: ");
                contractId = System.Console.ReadLine().Trim();
                if (contractId.ToLower() == "q")
                {
                    docontinue = false;
                    break;
                }

                System.Console.Write("Num Policies: ");
                numPolicies = System.Console.ReadLine().Trim();
                if (contractId.ToLower() == "q")
                {
                    docontinue = false;
                    break;
                }
                obj.RaiseRecordRequested(clientId,taxId, contractId,Int32.Parse(numPolicies));    
            }while(docontinue);

        }

        static void co_Requested(string clientId, string taxId, string contractId, int contractCount)
        {
            System.Console.WriteLine("\t[co_Requested] got event: clientId=" + clientId + " taxId="+taxId+" contractId=" + contractId );
        }
    }
}
