using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YetAnotherNewItemAlertWindowOfOutlook
{
    public class Column
    {
        private double? width = null;
        private string name = "";

        public string Name { get => name; set => name = value; }
        public double? Width { get => width; set => width = value; }
    }
}
