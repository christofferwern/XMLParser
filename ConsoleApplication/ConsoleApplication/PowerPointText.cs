using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class PowerPointText
    {
        private int _fontColor, _fontSize, _x, _y, _cx, _cy, _rotation;
        private string _type, _font, _alignment, _anchor;
        private Boolean _bold, _italic, _underline;

        public PowerPointText()
        {
            _fontColor = 0;
            _fontSize = 0;
            _x = 0;
            _y = 0;
            _cx = 0;
            _cy = 0;
            _rotation = 0;
            _type = "";
            _font = "";
            _alignment = "";
            _anchor = "";
            _bold = false;
            _italic = false;
            _underline = false;
        }

        public String toString()
        {
            return "Type: " + _type + "\n" + 
                   "  Font:        " + _font + "\n" +
                   "  Font size:   " + _fontSize + "\n" +
                   "  Font color:  " + _fontColor.ToString("X") + "\n" +
                   "  Size:        (" + _x + "," + _y + ")\n" +
                   "  Position:    (" + _cx + "," + _cy + ")\n" +
                   "  Anchor:      " + _anchor + "\n" +
                   "  Alignment:   " + _alignment + "\n" + 
                   "  B U I:       (" + _bold + ", " + _underline + ", " + _italic + ") \n";
        }

        public Boolean Underline
        {
            get { return _underline; }
            set { _underline = value; }
        }

        public Boolean Italic
        {
            get { return _italic; }
            set { _italic = value; }
        }

        public Boolean Bold
        {
            get { return _bold; }
            set { _bold = value; }
        }

        public string Anchor
        {
            get { return _anchor; }
            set { _anchor = value; }
        }

        public string Alignment
        {
            get { return _alignment; }
            set { _alignment = value; }
        }

        public string Font
        {
            get { return _font; }
            set { _font = value; }
        }

        public string Type
        {
            get { return _type; }
            set { _type = value; }
        }
        
        public int Rotation
        {
            get { return _rotation; }
            set { _rotation = value; }
        }

        public int Cy
        {
            get { return _cy; }
            set { _cy = value; }
        }

        public int Cx
        {
            get { return _cx; }
            set { _cx = value; }
        }

        public int Y
        {
            get { return _y; }
            set { _y = value; }
        }

        public int X
        {
            get { return _x; }
            set { _x = value; }
        }

        public int FontSize
        {
            get { return _fontSize; }
            set { _fontSize = value; }
        }

        public int FontColor
        {
            get { return _fontColor; }
            set { _fontColor = value; }
        }
    }
}
