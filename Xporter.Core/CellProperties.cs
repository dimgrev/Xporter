using System;
using System.Collections.Generic;
using System.Text;

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
