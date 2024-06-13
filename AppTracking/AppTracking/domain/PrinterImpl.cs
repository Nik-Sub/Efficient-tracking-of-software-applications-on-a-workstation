using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppTracking.domain
{
    interface PrinterImpl
    {
        void printReport(List<Dictionary<string, string>> apps);
    }
}
