using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HAT_API_PSADS
{
    class Program
    {
        static void Main(string[] args)
        {
            PsadsFlow flow = new PsadsFlow();
            flow.DoPsadsFlow();
        }
    }
}
