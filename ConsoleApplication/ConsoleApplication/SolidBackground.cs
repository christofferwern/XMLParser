using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class SolidBackground
    {
        private string _bgColor, _backgroundType, _fillType;

        
        private int _alpha;
        
        public SolidBackground()
        {
            _bgColor = "ffffff";
            _alpha = 0;
            _fillType = "solid";
            _backgroundType = "linear";
        }

        public string BgColor
        {
            get { return _bgColor; }
            set { _bgColor = value; }
        }

        public int Alpha
        {
            get { return _alpha; }
            set { _alpha = value; }
        }
        public string BackgroundType
        {
            get { return _backgroundType; }
            set { _backgroundType = value; }
        }

        public string toString()
        {
            return "BackgroundColor: " + _bgColor + "\nAlpha: " + _alpha + "\nType: " + _backgroundType;
        }

        public string FillType
        { 
            get{ return _fillType; }
            set { _fillType = value; } 
        }
    }
}
