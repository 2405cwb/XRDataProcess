using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNRoadFormatConverter.Entitys.Excel
{
    public class IRI 
    {
        public string RoadNumber { get; set; }

        public int SMile { get; set; }

        public int EMile { get; set; }

        public string RoadLane { get; set; }
        

        public double LeftIRI { get; set; }
        public double RightIRI { get; set; }

        public double JudgeIRI { get; set; }

        public double RQIValue { get; set; }

        public string Evaluation { get; set; }

        public string RoadType { get; set; }

        public double Speed { get; set; }

        public string Mark { get; set; }
    }
}
