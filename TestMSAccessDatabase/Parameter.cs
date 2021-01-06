using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestMSAccessDatabase
{
    class Parameter
    {
        public string name { get; set; }
        public object value { get; set; }

        public Parameter(string name, object value)
        {
            this.name = name;
            this.value = value;
        }
    }
}
