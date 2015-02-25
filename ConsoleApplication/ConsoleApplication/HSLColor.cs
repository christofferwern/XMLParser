using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class HSLColor
    {

        double _hue, _saturation, _luminance;
        public HSLColor()
        {
            _hue = 0;
            _saturation = 0;
            _luminance = 0;
        }
        public double Hue
        {
            get { return _hue; }
            set
            {
                _hue = value;
                _hue = _hue > 1 ? 1 : _hue < 0 ? 0 : _hue;
            }
        }

        public double Saturation
        {
            get { return _saturation; }
            set
            {
                _saturation = value;
                _saturation = _saturation > 1 ? 1 : _saturation < 0 ? 0 : _saturation;
            }
        }

        public double Luminance
        {
            get { return _luminance; }
            set
            {
                _luminance = value;
                _luminance = _luminance > 1 ? 1 : _luminance < 0 ? 0 : _luminance;
            }
        }

        public string toString()
        {
            return "\n H: " + Hue + "\n S: " + Saturation + "\n L: " + Luminance + "\n";
        }

    }
}
