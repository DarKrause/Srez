using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Srez.Models
{
    internal class Sale
    {
        public string dateSale { get; set; }
        public virtual Client client { get; set; }
    }
}
