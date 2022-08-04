using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace API_test
{
    public class MRTLevel
    {
        /*public MRTLevel(string level_name, int _index, double _decoration_THK, double _structure_THK, double _ceiling_THK, double _v_clearance)
        {
            name = level_name;
            index = _index;
            height = 0;
            selected = false;

            decoration_THK = _decoration_THK;
            structure_THK = _structure_THK;
            ceiling_THK = _ceiling_THK;
            v_clearance = _v_clearance;
        }*/
        public string name { get; set; }
        public int index { get; set; }
        public double height { get; set; }
        public bool selected { get; set; }

        public double floor_height { get; set; }
        public double display_height { get; set; }

        public double decoration_THK { get; set; }
        public double structure_THK { get; set; }
        public double ceiling_THK { get; set; }
        public double v_clearance { get; set; }

        public override string ToString()
        {
            return name;
        }

    }
}
