using System.Collections.Generic;

namespace Xporter
{
    public sealed class CellProperties : Dictionary<string, string>
    {
        public CellProperties() : base()
        {
        }
        public new void Add(string cell, string value)
        {
            base.Add(cell, value);
        }
    }
}
