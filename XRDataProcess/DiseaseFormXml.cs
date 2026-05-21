using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XRDataProcess
{
    public class DiseaseFormXml
    {
        public DiseaseFormXml(string roadGrad, string roadType, string diseaseName, string weight, string number)
        {
            RoadGrad = roadGrad;
            RoadType = roadType;
            DiseaseName = diseaseName;
            Weight = weight;
            Number = number;
        }


        public string RoadGrad { get; set; }
        public string RoadType { get; set; }

        public string DiseaseName { get; set; }

        public string Weight { get; set; }

        public string Number { get; set; }
    }
}
