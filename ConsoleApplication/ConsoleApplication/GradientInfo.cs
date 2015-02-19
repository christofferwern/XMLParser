using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Media.Converters;

namespace ConsoleApplication
{
    public class GradientInfo
    {
        private string _gradColor;
        private int _position;
        private float _alpha, _tint, _saturationMod, _shade, _luminance;

        private Color _color;

        public GradientInfo()
        {
            _gradColor = "ffffff";
            _position = 0;
            _alpha = 1;
            _tint = 100;
            _saturationMod = 0;
            _shade = 0;
            _color = new Color();
        }

        public void convert()
        {
            System.Drawing.Color col = System.Drawing.ColorTranslator.FromHtml("#FFFF00");
            /*float tint = 0.5f;

            Console.WriteLine(col.GetBrightness());
            

            int coll = int.Parse(col.R.ToString(), System.Globalization.NumberStyles.HexNumber);
            int convert = (int)(col.R * tint);
            string hex = convert.ToString("X");

            int collG = int.Parse(col.G.ToString(), System.Globalization.NumberStyles.HexNumber);
            int convertG = (int)(col.G * tint);
            string hexG = convertG.ToString("X");

            int collB = int.Parse(col.B.ToString(), System.Globalization.NumberStyles.HexNumber);
            int convertB = (int)(col.B * tint);
            string hexB = convertB.ToString("X");*/

            /*Console.WriteLine("Color: " + col +
                               "\n R: " + col.R + ", conR:" + convert +
                                "\n G: " + col.G + ", conG:" + convertG +
                                "\n B: " + col.B + ", conB:" + convertB );*/
            //Console.WriteLine("new Color: " + hex + ":" + hexG + ":"+ hexB);


            if(_tint!=100)
                Console.WriteLine("R: " + col.R * (_tint / 100) + ", G: " + col.G * (_tint / 100) + ", B: " + col.B * (_tint / 100));
            //Console.WriteLine("H: " + col.GetHue() + ", S: " + col.GetSaturation() + ", B: " + col.GetBrightness());

        }


        public Color Color
        {
            get { return _color; }
            set { _color = value; }
        }


        public float Luminance
        {
            get { return _luminance; }
            set { _luminance = value; }
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
        public string GradColor
        {
            get { return _gradColor; }
            set 
            { 
                _gradColor = value;
                _color = ColorTranslator.FromHtml("#" + _gradColor);
            }
        }

        public float SaturationMod
        {
            get { return _saturationMod; }
            set { _saturationMod = value; }
        }

        public float Shade
        {
            get { return _shade; }
            set { _shade = value; }
        }

        public float Tint
        {
            get { return _tint; }
            set { _tint = value; }
        }

        public string toString()
        {
            return "Color: " + _color + "\nPosition: " + _position + "\nAlpha: " + _alpha +
                    "\nTint: " + _tint + "\nSatMode: " + _saturationMod + "\nShade: " + _shade + "\n";
        }
    }
}
