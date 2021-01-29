using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XporterConsole
{
    internal class Students
    {
        public Students()
        {
        }

        public int ID { get; set; }
        public List<string> FirstName { get; set; }
        public List<string> LastName { get; set; }

        //public override string ToString() => FirstName + "  " + ID;
    }
}
