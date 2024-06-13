using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppTracking.domain
{
    class Printer{

        PrinterImpl printerImpl;


        public Printer(PrinterImpl printerImplementation)
        {
            this.printerImpl = printerImplementation;
        }
        public void printReport(List<Dictionary<string, string>> apps)
        {
            printerImpl.printReport(apps);
        }
    }
}
