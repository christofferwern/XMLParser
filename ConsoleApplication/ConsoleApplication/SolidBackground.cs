using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class SolidBackground
    {
        private string _bgColor, _backgroundType, _fillType;
        private int _alpha, _tint, _luminanceMod, _saturationMod;

        private Color _color;

        public SolidBackground()
        {
            _bgColor = "ffffff";
            _alpha = 0;
            _fillType = "solid";
            _backgroundType = "linear";
            _color = new Color();
        }

        public string BgColor
        {
            get { return _bgColor; }
            set 
            { 
                _bgColor = value;
                _color = ColorTranslator.FromHtml("#"+_bgColor);
            }
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

        public int SaturationMod
        {
            get { return _saturationMod; }
            set { _saturationMod = value; }
        }

        public int LuminanceMod
        {
            get { return _luminanceMod; }
            set { _luminanceMod = value; }
        }

        public int Tint
        {
            get { return _tint; }
            set { _tint = value; }
        }
        public Color Color
        {
            get { return _color; }
            set { _color = value; }
        }

        public string FillType
        { 
            get{ return _fillType; }
            set { _fillType = value; } 
        }

        public string toString()
        {
            return "BackgroundColor: " + _bgColor + "\nAlpha: " + _alpha + "\nType: " + _backgroundType +
                "\nTint: " + _tint + "\nLumMod: " + _luminanceMod + "\nSatMod: " + _saturationMod + "\n";
        }
    }
}
