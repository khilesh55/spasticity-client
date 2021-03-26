using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpasticityClient
{
    public class SessionData
    {
        public long TimeStamp { get; set; }

        public float AngVelX_A { get; set; }
        public float AngVelY_A { get; set; }
        public float AngVelZ_A { get; set; }

        public float OrientX_A { get; set; }
        public float OrientY_A { get; set; }
        public float OrientZ_A { get; set; }

        public float AngVelX_B { get; set; }
        public float AngVelY_B { get; set; }
        public float AngVelZ_B { get; set; }

        public float OrientX_B { get; set; }
        public float OrientY_B { get; set; }
        public float OrientZ_B { get; set; }

        public float EMG { get; set; }
        public float Force { get; set; }
    }
}
