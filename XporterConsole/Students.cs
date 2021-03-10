using System.Collections.Generic;

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
