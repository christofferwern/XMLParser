using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class TextFragment
    {
        private String _text;
        private int _x, _y, _styleId;
        private XmlDocument _rootOfDocument;

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

        public XmlElement getFragmentChild()
        {
            XmlElement f = _rootOfDocument.CreateElement("f");

            XmlAttribute t_attr = _rootOfDocument.CreateAttribute("t");
            t_attr.Value = _text;

            XmlAttribute x_attr = _rootOfDocument.CreateAttribute("x");
            x_attr.Value = _x.ToString();

            XmlAttribute y_attr = _rootOfDocument.CreateAttribute("y");
            y_attr.Value = _y.ToString();

            XmlAttribute s_attr = _rootOfDocument.CreateAttribute("s");
            s_attr.Value = _styleId.ToString();

            f.Attributes.Append(t_attr);
            f.Attributes.Append(x_attr);
            f.Attributes.Append(y_attr);
            f.Attributes.Append(s_attr);

            return f;
        }

        public void setXMLDocumentRoot(ref XmlDocument rootOfDocument)
        {
            _rootOfDocument = rootOfDocument;
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