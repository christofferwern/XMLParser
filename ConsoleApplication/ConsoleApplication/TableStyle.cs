using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class TableStyle
    {
        int _lineSize, _fontSize, _fillAlpha;
        string _fontColor, _lineColor, _fillColor, _type;

        public TableStyle()
        {
            _lineSize = 0;
            _fontSize = 0;
            _fillAlpha = 0;

            _fontColor = "";
            _lineColor = "";
            _fillColor = "";
            _type = "";
        }

        public override string ToString()
        {
 	        return "Type: " + _type + "\n" +
                   "Fill: " + _fillColor + ", " + _fillAlpha + "\n" +
                   "Line: " + _lineColor + ", " + _lineSize + "\n" +
                   "Font: " + _fontColor + ", " + _fontSize;
        }

        public string Type
        {
            get { return _type; }
            set { _type = value; }
        }

        public string FillColor
        {
            get { return _fillColor; }
            set { _fillColor = value; }
        }

        public string LineColor
        {
            get { return _lineColor; }
            set { _lineColor = value; }
        }

        public string FontColor
        {
            get { return _fontColor; }
            set { _fontColor = value; }
        }
        public int FillAlpha
        {
            get { return _fillAlpha; }
            set { _fillAlpha = value; }
        }

        public int FontSize
        {
            get { return _fontSize; }
            set { _fontSize = value; }
        }

        public int LineSize
        {
            get { return _lineSize; }
            set { _lineSize = value; }
        }
        

        
    }
}
