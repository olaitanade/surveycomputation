using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiverrExcelForms
{
    class ClosedPt
    {
        //Input variables
        public string Standpt { get; set; }
        public string TargetPoint { get; set; }
        public double Degree { get; set; }
        public double Minute{ get; set; }
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
        public double StartZ { get; set; }
        public double E { get; set; }
        public double N { get; set; }
        public double Sum_distance{ get; set; }
        public double Diff_x { get; set; }
        public double Diff_y { get; set; }
        public double Diff_z { get; set; }
    }
}
