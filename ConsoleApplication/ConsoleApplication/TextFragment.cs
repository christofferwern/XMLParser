using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class TextFragment
    {
        private String _text;
        private int _x, _y, _styleId;

        public int StyleId
        {
            get { return _styleId; }
            set { _styleId = value; }
        }

        public TextFragment()
        {
            _text = "text";
            _x = 0;
            _y = 0;
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


    }
}