using System;
using System.Collections.Generic;
using System.Linq;
using UnifiedOpennessLibrary;

namespace UnifiedOpennessLibraryDemoSample
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var unifiedData = new UnifiedOpennessConnector("V18", args, new List<CmdArgument>(), "DemoSample"))
            {
                Work(unifiedData);
            }
        }
        static void Work(UnifiedOpennessConnector unifiedData)
        {

            using (var transaction = unifiedData.AccessObject.Transaction(unifiedData.TiaPortalProject, "DemoSample - revert my fancy changes"))
            {
                Console.WriteLine(unifiedData.DeviceName);
                unifiedData.Screens.ToList()[0].Name = "FOO";
                transaction.CommitOnDispose();
            }
        }
    }
}
