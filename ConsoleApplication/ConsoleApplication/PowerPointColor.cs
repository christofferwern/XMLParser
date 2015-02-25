using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class PowerPointColor
    {
        private int _tint,
                    _shade,
                    _comp,
                    _inv,
                    _gray,
                    _alpha,
                    _alphaOff,
                    _alphaMod,
                    _hueOff,
                    _hueMod,
                    _sat,
                    _satOff,
                    _satMod,
                    _lum,
                    _lumOff,
                    _lumMod,
                    _red,
                    _redOff,
                    _redMod,
                    _green,
                    _greenOff,
                    _greenMod,
                    _blue,
                    _blueOff,
                    _blueMod,
                    _gamma,
                    _invGamma;

    
        private string _color;

        public PowerPointColor()
        {
            _color = "";
            _tint = 0;
            _shade = 0;
            _comp = 0;
            _inv = 0;
            _gray = 0;
            _alpha = 100000;
            _alphaOff = 0;
            _alphaMod = 0;
            _hueOff = 0;
            _hueMod = 0;
            _sat = 0;
            _satOff = 0;
            _satMod = 0;
            _lum = 0;
            _lumOff = 0;
            _lumMod = 0;
            _red = 0;
            _redOff = 0;
            _redMod = 0;
            _green = 0;
            _greenOff = 0;
            _greenMod = 0;
            _blue = 0;
            _blueOff = 0;
            _blueMod = 0;
            _gamma = 0;
            _invGamma = 0;
        }

        //TODO
        public string getAdjustedColor()
        {


            return _color;
        }

        public override string ToString()
        {
 	        string output =  "PowerPointColor\n";

            output +=                       " color:    " + _color + "\n";
            output += (_tint != 0) ?        " tint:     " + _tint + "\n" : "";
            output += (_shade != 0) ?       " shade:    " + _shade + "\n" : "";
            output += (_inv != 0) ?         " inv:      " + _inv + "\n" : "";
            output += (_gray != 0) ?        " gray:     " + _gray + "\n" : "";
            output += (_alpha != 0) ?       " alpha:    " + _alpha + "\n" : "";
            output += (_alphaOff != 0) ?    " alphaOff: " + _alphaOff + "\n" : "";
            output += (_alphaMod != 0) ?    " alphaMod: " + _alphaMod + "\n" : "";
            output += (_hueOff != 0) ?      " hueOff:   " + _hueOff + "\n" : "";
            output += (_hueMod != 0) ?      " hueMod:   " + _hueMod + "\n" : "";
            output += (_satOff != 0) ?      " satOff:   " + _satOff + "\n" : "";
            output += (_satMod != 0) ?      " satMod:   " + _satMod + "\n" : "";
            output += (_lum != 0) ?         " lum:      " + _lum + "\n" : "";
            output += (_lumOff != 0) ?      " lumOff:   " + _lumOff + "\n" : "";
            output += (_lumMod != 0) ?      " lumMod:   " + _lumMod + "\n" : "";
            output += (_red != 0) ?         " red:      " + _red + "\n" : "";
            output += (_redOff != 0) ?      " redOff:   " + _redOff + "\n" : "";
            output += (_redMod != 0) ?      " redMod:   " + _redMod + "\n" : "";
            output += (_green != 0) ?       " green:    " + _green + "\n" : "";
            output += (_greenOff != 0) ?    " greenOff: " + _greenOff + "\n" : "";
            output += (_greenMod != 0) ?    " greenMod: " + _greenMod + "\n" : "";
            output += (_blue != 0) ?        " blue:     " + _blue + "\n" : "";
            output += (_blueOff != 0) ?     " blueOff:  " + _blueOff + "\n" : "";
            output += (_blueMod != 0) ?     " blueMod:  " + _blueMod + "\n" : "";
            output += (_gamma != 0) ?       " gamma:    " + _gamma + "\n" : "";
            output += (_invGamma != 0) ?    " invGamma: " + _invGamma + "\n": "";         

            return output;
        }

        public int SatMod
        {
            get { return _satMod; }
            set { _satMod = value; }
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

        public int LumMod
        {
            get { return _lumMod; }
            set { _lumMod = value; }
        }

        public int Shade
        {
            get { return _shade; }
            set { _shade = value; }
        }

        public int InvGamma
        {
            get { return _invGamma; }
            set { _invGamma = value; }
        }

        public int Gamma
        {
            get { return _gamma; }
            set { _gamma = value; }
        }

        public int BlueMod
        {
            get { return _blueMod; }
            set { _blueMod = value; }
        }

        public int BlueOff
        {
            get { return _blueOff; }
            set { _blueOff = value; }
        }

        public int Blue
        {
            get { return _blue; }
            set { _blue = value; }
        }

        public int GreenMod
        {
            get { return _greenMod; }
            set { _greenMod = value; }
        }

        public int GreenOff
        {
            get { return _greenOff; }
            set { _greenOff = value; }
        }

        public int Green
        {
            get { return _green; }
            set { _green = value; }
        }

        public int RedMod
        {
            get { return _redMod; }
            set { _redMod = value; }
        }

        public int RedOff
        {
            get { return _redOff; }
            set { _redOff = value; }
        }

        public int Red
        {
            get { return _red; }
            set { _red = value; }
        }

        public int LumOff
        {
            get { return _lumOff; }
            set { _lumOff = value; }
        }

        public int Lum
        {
            get { return _lum; }
            set { _lum = value; }
        }

        public int SatOff
        {
            get { return _satOff; }
            set { _satOff = value; }
        }

        public int Sat
        {
            get { return _sat; }
            set { _sat = value; }
        }

        public int HueOff
        {
            get { return _hueOff; }
            set { _hueOff = value; }
        }

        public int AlphaMod
        {
            get { return _alphaMod; }
            set { _alphaMod = value; }
        }

        public int AlphaOff
        {
            get { return _alphaOff; }
            set { _alphaOff = value; }
        }

        public int Gray
        {
            get { return _gray; }
            set { _gray = value; }
        }

        public int Inv
        {
            get { return _inv; }
            set { _inv = value; }
        }

        public int Comp
        {
            get { return _comp; }
            set { _comp = value; }
        }


        public int HueMod
        {
            get { return _hueMod; }
            set { _hueMod = value; }
        }

    }
}
