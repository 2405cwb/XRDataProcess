using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNRoadFormatConverter.Entitys.Excel
{
    public class PCI 
    {
        public string RoadNumber { get; set; }

        public int SMile { get; set; }

        public int EMile { get; set; }

        public string RoadLane { get; set; }
        

        public double DR { get; set; }
        public double PCIValue { get; set; } 

        public string Evaluation { get; set; }

        public string RoadType { get; set; }

        public double Speed { get; set; }

        public string Mark { get; set; }
    }
}
