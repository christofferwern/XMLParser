using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class GradientInfo
    {
        private string _color;
        private int _position;
        private float _alpha;
        
        public GradientInfo()
        {
            _color = "ffffff";
            _position = 0;
            _alpha = 0;
        }
        public float Alpha
        {
            get { return _alpha; }
            set { _alpha = value; }
        }
        public int Position
        {
            get { return _position; }
            set { _position = value; }
        }
        public string Color
        {
            get { return _color; }
            set { _color = value; }
        }

        public string toString()
        {
            return "Color: " + _color + ", Position: " + _position + ", Alpha: " + _alpha;
        }
    }
}
