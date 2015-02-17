using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class TextStyle
    {
        private Boolean _bold, _underline, _italic;
        private int _fontSize;
        private string _font, _fontColor;
        private XmlDocument _rootOfDocument;

        public TextStyle()
        {
            _bold = false;
            _underline = false;
            _italic = false;
            _fontSize = 0;
            _fontColor = "000000";
            _font = "Default";
        }

        public string attrubiteValue()
        {
            string style = "k";

            if(_bold)
                style += "b";
            if(_underline)
                style += "u";
            if(_italic)
                style += "i";

            return _font + "," + _fontSize.ToString() + "," + _fontColor.ToString() + "," + style;
        }

        private XmlAttribute getStylesAttributes()
        {
            XmlAttribute attrStyle = _rootOfDocument.CreateAttribute("style");
            attrStyle.Value = attrubiteValue();

            return attrStyle;
        }

        public XmlElement getStylesChild()
        {
            XmlElement child = _rootOfDocument.CreateElement("s");
            child.Attributes.Append(getStylesAttributes());

            return child;
        }

        public XmlElement getFontNode()
        {
            XmlElement FONT = _rootOfDocument.CreateElement("FONT");

            return FONT;

        }

        public string getTextNode(string text)
        {

            string temp = text;

            if (Bold)
                temp = "<B>" + text + "</B>";
            if(Italic)
                temp = "<I>" + text + "</I>";
            if(Underline)
                temp = "<U>" + text + "</U>";

            return temp;

        }

        public void setXMLDocumentRoot(ref XmlDocument rootOfDocument)
        {
            _rootOfDocument = rootOfDocument;
        }

        public string Font
        {
            get { return _font; }
            set { _font = value; }
        }

        public string FontColor
        {
            get { return _fontColor; }
            set { _fontColor = value; }
        }

        public int FontColorInteger()
        {
            return int.Parse(_fontColor, System.Globalization.NumberStyles.HexNumber);
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

        public bool isEqual(object obj)
        {
            TextStyle other = obj as TextStyle;

            if (other == null)
                return false;

            if ((Font == other.Font) && (FontColor == other.FontColor) && (FontSize == other.FontSize) && (Italic == other.Italic) && (Underline == other.Underline) && (Bold == other.Bold))
            {
                return true; 
            }
            else
            {
                return false;
            }

          /*  return (Font == other.Font)
                && (FontColor == other.FontColor)
                && (FontSize == other.FontSize)
                && (Italic == other.Italic)
                && (Underline == other.Underline)
                && (Bold == other.Bold);*/
        }

        //public static bool operator ==(TextStyle x, TextStyle y)
        //{
        //    if (x == null || y == null)
        //        return false;

        //     return      (x.Font == y.Font)
        //              && (x.FontColor == y.FontColor)
        //              && (x.FontSize == y.FontSize)
        //              && (x.Italic == y.Italic)
        //              && (x.Underline == y.Underline)
        //              && (x.Bold == y.Bold);
        //}

        //public static bool operator !=(TextStyle x, TextStyle y)
        //{
        //    if (x == null || y == null)
        //        return false;

        //    return !(x == y);
        //}

        public string toString()
        {
            return "Font:   " + _font + "\n" +
                   "Size:   " + _fontSize + "\n" +
                   "Color:  " + _fontColor + "\n" +
                   "B U I:  (" + _bold + ", " + _underline + ", " + _italic + ") \n";
        }

    }
}
