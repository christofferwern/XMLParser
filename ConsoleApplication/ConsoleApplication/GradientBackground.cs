using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class GradientBackground
    {
        //pos, color, alpha
        private List<GradientInfo> _gradientList;
        private string _bgType, _gradientType;


        private int _angle;
        public GradientBackground()
        {

            _gradientList = new List<GradientInfo>();
            _angle = 0;
            _bgType = "gradient";
            _gradientType = "";
        }

        public float getAlpha1()
        {
            return _gradientList.First().Alpha;
        }

        public int Angle
        {
            get { return _angle; }
            set { _angle = value; }
        }

        public List<GradientInfo> GradientList
        {
            get { return _gradientList; }
            set { _gradientList = value; }
        }

        public string GradientType
        {
            get { return _gradientType; }
            set { _gradientType = value; }
        }

        public string BgType
        {
            get { return _bgType; }
            set { _bgType = value; }
        }

        public string toString()
        {
            string output = "Infolist:\n";
            foreach (var item in _gradientList)
            {
                output += item.toString() + "\n";
            }
            output += "GradientType: " + GradientType + "\n";

            return output;
        }
    }
}
