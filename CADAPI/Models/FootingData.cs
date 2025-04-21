using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CADAPI.Models
{
    public class FootingData
    {
        public double X, Y;
        public double WidthPC, LengthPC, DepthPC;
        public double WidthRC, LengthRC, DepthRC;
        public string Tag { get; set; }
    }
}
