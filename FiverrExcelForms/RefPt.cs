using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiverrExcelForms
{
    class RefPt
    {
        //Input variables
        public ClosedPt RefPoint { get; set; }
        public string TargetPoint { get; set; }
        public double Degree { get; set; }
        public double Minute { get; set; }
        public double Second { get; set; }
        public double Distance { get; set; }
        public double DeltaZ { get; set; }

        //Output variable
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public double Cw { get; set; }
        public double Ccw { get; set; }
        public double WorldA { get; set; }
        public double E { get; set; }
        public double N { get; set; }
        public double Diff_z { get; set; }
    }
}
