using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace API_test
{
    public class MRTStation
    {
        public MRTStation()
        {

        }

        public double length { get; set; }
        public double width { get; set; }
        public double left { get; set; }
        public double right { get; set; }
        public double top { get; set; }
        public double bottom { get; set; }

        public double centerX { get; set; }
        public double centerY { get; set; }

        public double c_wall { get; set; }
        public double s_wall { get; set; }
    }
}
