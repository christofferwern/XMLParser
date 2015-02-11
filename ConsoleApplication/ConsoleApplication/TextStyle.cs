using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class TextStyle
    {
        private Boolean _bold, _underline, _italic;
        private int _fontSize, _fontColor;
        private string _font;

        public TextStyle()
        {
            _bold = false;
            _underline = false;
            _italic = false;
            _fontSize = 0;
            _fontColor = Color.Black.ToArgb();
            _font = "Default";
        }

        public String toString()
        {
            return "Font:   " + _font + "\n" +
                   "Size:   " + _fontSize + "\n" +
                   "Color:  " + _fontColor + "\n" +
                   "B U I:  (" + _bold + ", " + _underline + ", " + _italic + ") \n";
        }

        public string Font
        {
            get { return _font; }
            set { _font = value; }
        }

        public int FontColor
        {
            get { return _fontColor; }
            set { _fontColor = value; }
        }

        public int FontSize
        {
            get { return _fontSize; }
            set { _fontSize = value; }
        }

        public Boolean Italic
        {
            get { return _italic; }
            set { _italic = value; }
        }

        public Boolean Underline
        {
            get { return _underline; }
            set { _underline = value; }
        }

        public Boolean Bold
        {
            get { return _bold; }
            set { _bold = value; }
        }
    }
}
