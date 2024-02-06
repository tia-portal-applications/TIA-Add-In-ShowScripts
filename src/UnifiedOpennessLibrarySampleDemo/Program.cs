using System;
using System.Collections.Generic;
using UnifiedOpennessLibrary;

namespace UnifiedOpennessLibraryDemoSample
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var unifiedData = new UnifiedOpennessConnector("V18", args, new List<CmdArgument>(), "DemoSample"))
            {
                Console.WriteLine(unifiedData.DeviceName);
            }
        }
    }
}
