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
        Boolean _bold, _underline, _italic;
        string _font;
        int _fontColor, _fontSize;

        public TextStyle()
        {
            _bold = false;
            _underline = false;
            _italic = false;

            _font = "Arial";
            _fontColor = Color.Black.ToArgb();
            _fontSize = 14;
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

        public string Font
        {
            get { return _font; }
            set { _font = value; }
        }
    }
}
