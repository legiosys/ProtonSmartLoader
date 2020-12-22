using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Proton_Loader.Models
{
    public class Template
    {
        public string Name { get; set; }
        public string Id { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
