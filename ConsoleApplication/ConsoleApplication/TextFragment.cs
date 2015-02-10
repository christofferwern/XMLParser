using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class TextFragment
    {
        private int _x, _y;
        private String _text;
        private TextStyle _textStyle;

        public TextFragment()
        {
            _x = 0;
            _y = 0;
            _text = "Text";
            _textStyle = new TextStyle();
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

        public String Text
        {
            get { return _text; }
            set { _text = value; }
        }

        internal TextStyle TextStyle
        {
            get { return _textStyle; }
            set { _textStyle = value; }
        }

    }
}
