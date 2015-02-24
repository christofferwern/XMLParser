using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class PowerPointColor
    {
        private int _shade, _tint, _luminance, _alpha, _saturation;
        private string _color;

        public PowerPointColor()
        {
            _color = "";
            _shade = 0;
            _tint = 0;
            _luminance = 0;
            _alpha = 100000;
            _saturation = 0;
        }

        public override string ToString()
        {
 	        string output =  "PowerPointColor\n";

            output +=   "  Color:      " + _color + "\n" +
                        "  Shade:      " + _shade + "\n" +
                        "  Tint:       " + _tint + "\n" +
                        "  Luminance:  " + _luminance + "\n" +
                        "  Alpha:      " + _alpha + "\n" +
                        "  Saturation: " + _saturation;

            return output;
        }

        public int Saturation
        {
            get { return _saturation; }
            set { _saturation = value; }
        }

        public string Color
        {
            get { return _color; }
            set { _color = value; }
        }

        public int Alpha
        {
            get { return _alpha; }
            set { _alpha = value; }
        }

        public int Tint
        {
            get { return _tint; }
            set { _tint = value; }
        }

        public int Luminance
        {
            get { return _luminance; }
            set { _luminance = value; }
        }

        public int Shade
        {
            get { return _shade; }
            set { _shade = value; }
        }
    }
}
