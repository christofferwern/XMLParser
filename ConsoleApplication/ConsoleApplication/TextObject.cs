using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class TextObject : SceneObjectDecorator
    {
        private string _align, _antiAlias, _font, _autosize, _color;
        private List<TextStyle> _styleList;
        private List<TextFragment> _fragmentsList;
        private int _leading, _letterSpacing, _size;
        private Boolean _bold, _italic, _underline, _selectable, _runningText, _useScroller;
        private XmlDocument _doc;
        private string[] _textObjectPropertiesAttributes = new string[14]{   "font", "align", "color", "italic", "bold", "underline", "size", "runningText",
                                                                             "autosize", "leading", "letterSpacing", "antiAlias", "useScroller", "selectable"};
        private SimpleSceneObject _simpleSceneObject;

        public TextObject(SimpleSceneObject sceneobject) : base(sceneobject) 
        {
            _simpleSceneObject = sceneobject;
                       
            _align = "left";
            _antiAlias = "normal";
            _font = "Arial";
            _autosize = "none";
            _color = "000000";
            _leading = 0;
            _letterSpacing = 0;
            _size = 20;
            _bold = false;
            _italic = false;
            _underline = false;
            _selectable = false;
            _runningText = true;
            _useScroller = false;

            string objectType = "com.customObjects.TextObject";
            _styleList = new List<TextStyle>();
            _fragmentsList = new List<TextFragment>();

            setProperties(new Properties(true, false, true, true, true, true, true, true,
                                         false, true, true, true, true, true, true));

            sceneobject.setObjectType(objectType);
        }

        public override object Clone()
        {
            SimpleSceneObject simpleSceneObject = _simpleSceneObject;

            TextObject textObject = new TextObject(simpleSceneObject);

            textObject.Align = _align;
            textObject.AntiAlias = _antiAlias;
            textObject.Font = _font;
            textObject.Autosize = _autosize;
            textObject.Color = _color;
            textObject.Leading = _leading;
            textObject.LetterSpacing = _letterSpacing;
            textObject.Size = _size;
            textObject.Bold = _bold;
            textObject.Italic = _italic;
            textObject.Underline = _underline;
            textObject.Selectable = _selectable;
            textObject.RunningText = _runningText;
            textObject.UseScroller = _useScroller;

            string objectType = "com.customObjects.TextObject";
            _styleList = new List<TextStyle>();
            _fragmentsList = new List<TextFragment>();

            setProperties(new Properties(true, false, true, true, true, true, true, true,
                                         false, true, true, true, true, true, true));

            textObject.setObjectType(objectType);

            return textObject;
        }

        public TextObject setClipWidth(int clipWidth)
        {
            _simpleSceneObject.ClipWidth = clipWidth;
            return (TextObject)Clone();
        }

        public void addToStyleList(TextStyle textStyle)
        {
            bool isEqual = false;

            if (_styleList.Count == 0)
                _styleList.Add(textStyle);
            else
            {
                foreach (TextStyle item in _styleList)
                    if (textStyle.isEqual(item))
                    {
                        isEqual = true;
                        break;
                    }
                
                if (!isEqual)
                    _styleList.Add(textStyle);
            }
            
        }

        public XmlElement getTextObjectPropertiesNode()
        {
            XmlElement textObjectPropNode = getXMLDocumentRoot().CreateElement("textObjectProperties");

            //XmlElement stylesNode = getStylesNode();
            //XmlElement fragmentNode = getFragmentsNode();
            //XmlElement textNode = getTextNode();
            //textObjectPropNode.AppendChild(textNode);

            XmlElement textNode = getXMLDocumentRoot().CreateElement("text");
            textNode.InnerText = getHTML();
            textObjectPropNode.AppendChild(textNode);

            _doc = base.getXMLDocumentRoot();

            foreach (string s in _textObjectPropertiesAttributes)
            {
                XmlAttribute xmlAttr = _doc.CreateAttribute(s);

                FieldInfo fieldInfo = GetType().GetField("_" + s, BindingFlags.NonPublic | BindingFlags.Instance);
                if (fieldInfo!=null)
                    xmlAttr.Value = fieldInfo.GetValue(this).ToString().ToLower();

                textObjectPropNode.Attributes.Append(xmlAttr);
            }

            //textObjectPropNode.AppendChild(stylesNode);
            //textObjectPropNode.AppendChild(fragmentNode);

            return textObjectPropNode;
        }

        public XmlElement getPropertiesNode()
        {
            string properties = getProperties().toString();
            XmlElement prop = getXMLDocumentRoot().CreateElement("properties");
            prop.InnerText = properties;

            return prop;
        }

        public string getHTML()
        {
            string HTML = "";

            HTML += "<TEXTFORMAT LEFTMARGIN=\"1\" RIGHTMARGIN=\"2\">";

            TextStyle newStyle = new TextStyle(), oldStyle = new TextStyle();

            bool bold = false, underline = false, italic = false;
            int fontCount = 0;

            foreach (TextFragment textFragment in _fragmentsList)
            {
                if (textFragment.NewParagraph)
                    HTML += "<br>";

                newStyle = StyleList[textFragment.StyleId];

                //First fragment
                if (_fragmentsList.IndexOf(textFragment) == 0)
                {
                    HTML += "<P ALIGN=\"" + newStyle.Alignment + "\">";

                    fontCount++;
                    HTML += "<FONT FACE=\"" + newStyle.Font + "\" SIZE=\"" + newStyle.FontSize + "\" COLOR=\"#" + newStyle.FontColor + "\" LETTERSPACING=\"0\" KERNING=\"1\">";
                    
                    if (newStyle.Bold)
                    {
                        HTML += "<B>";
                        bold = true;
                    }
                    if (newStyle.Underline)
                    {
                        HTML += "<U>";
                        underline = true;
                    }
                    if (newStyle.Italic)
                    {
                        HTML += "<I>";
                        italic = true;
                    }
                     
                    for (int i = 0; i < textFragment.Level; i++)
                        HTML += "\t";

                    HTML += textFragment.Text.Replace("<", "&#60;").Replace(">", "&#62;");

                    oldStyle = newStyle;
                    continue;
                }

                if (oldStyle != newStyle)
                {
                    if (oldStyle.Font != newStyle.Font ||
                        oldStyle.FontColor != newStyle.FontColor ||
                        oldStyle.FontSize != newStyle.FontSize ||
                        oldStyle.Alignment != newStyle.Alignment)
                    {
                        if (oldStyle.Alignment != newStyle.Alignment)
                        {
                            for (int i = 0; i < fontCount; i++)
                                HTML += "</FONT>";

                            HTML += "</P>";
                            HTML += "</TEXTFORMAT>";
                            fontCount = 0;

                            HTML += "<P ALIGN=\"" + newStyle.Alignment + "\">";

                            fontCount++;
                            HTML += "<FONT FACE=\"" + newStyle.Font + "\" SIZE=\"" + newStyle.FontSize + "\" COLOR=\"#" + newStyle.FontColor + "\" LETTERSPACING=\"0\" KERNING=\"1\">";
                        }

                        HTML += (bold) ? "</B>" : "";
                        HTML += (underline) ? "</U>" : "";
                        HTML += (italic) ? "</I>" : "";

                        bold = false;
                        underline = false;
                        italic = false;

                        fontCount++;

                        HTML += "<FONT ";

                        if (oldStyle.Font != newStyle.Font)
                            HTML += "FACE=\"" + newStyle.Font + "\" ";
                        if (oldStyle.FontSize != newStyle.FontSize)
                            HTML += "SIZE=\"" + newStyle.FontSize + "\" ";
                        if (oldStyle.FontColor != newStyle.FontColor)
                            HTML += "COLOR=\"#" + newStyle.FontColor + "\" ";

                        HTML += ">";

                        if (newStyle.Bold)
                        {
                            HTML += "<B>";
                            bold = true;
                        }
                        if (newStyle.Underline)
                        {
                            HTML += "<U>";
                            underline = true;
                        }
                        if (newStyle.Italic)
                        {
                            HTML += "<I>";
                            italic = true;
                        }

                        for (int i = 0; i < textFragment.Level; i++)
                            HTML += "\t";

                        HTML += textFragment.Text.Replace("<", "&#60;").Replace(">", "&#62;");
                    }
                    else
                    {
                        if (newStyle.Bold != bold)
                        {
                            HTML += (newStyle.Bold) ? "<B>" : "</B>";
                            bold = (newStyle.Bold) ? true : false;
                        }
                        if (newStyle.Underline != underline)
                        {
                            HTML += (newStyle.Underline) ? "<U>" : "</U>";
                            underline = (newStyle.Underline) ? true : false;
                        }
                        if (newStyle.Italic != italic)
                        {
                            HTML += (newStyle.Italic) ? "<I>" : "</I>";
                            italic = (newStyle.Italic) ? true : false;
                        }

                        for (int i = 0; i < textFragment.Level; i++)
                            HTML += "\t";

                        HTML += textFragment.Text.Replace("<", "&#60;").Replace(">", "&#62;");
                    }
                }
                else
                {
                    for (int i = 0; i < textFragment.Level; i++)
                        HTML += "\t";

                    HTML += textFragment.Text.Replace("<", "&#60;").Replace(">", "&#62;");
                }

                for (int i = 0; i < textFragment.Breaks; i++)
                    HTML += "<br>";

                oldStyle = newStyle;
            }

            HTML += (bold) ? "</B>" : "";
            HTML += (underline) ? "</U>" : "";
            HTML += (italic) ? "</I>" : "";

            for (int i = 0; i < fontCount; i++)
                HTML += "</FONT>";

            HTML += "</P>";
            HTML += "</TEXTFORMAT>";

            return HTML;
        }

        public XmlElement getTextNode()
        {
            XmlElement textNode = getXMLDocumentRoot().CreateElement("text");
            XmlElement textFormatNode = getXMLDocumentRoot().CreateElement("TEXTFORMAT");
            XmlElement pNode = getPnode();

            textFormatNode.AppendChild(pNode);
            textNode.AppendChild(textFormatNode);

            return textNode;
        }

        public XmlElement getPnode()
        {
            XmlElement pNode = getXMLDocumentRoot().CreateElement("P");
            
            List<XmlElement> fontList = new List<XmlElement>();
            TextStyle old_style = new TextStyle();

            for (int i = 0; i < _fragmentsList.Count; i++)
            {
                XmlElement temp = getXMLDocumentRoot().CreateElement("FONT");
                TextStyle style = _styleList[_fragmentsList[i].StyleId];

                if (i == 0)
                {
                    XmlAttribute font = getXMLDocumentRoot().CreateAttribute("FACE");
                    font.Value = style.Font;
                    XmlAttribute size = getXMLDocumentRoot().CreateAttribute("SIZE");
                    size.Value = style.FontSize.ToString();
                    XmlAttribute color = getXMLDocumentRoot().CreateAttribute("COLOR");
                    color.Value = style.FontColor.ToString();
                    temp.Attributes.Append(font);
                    temp.Attributes.Append(size);
                    temp.Attributes.Append(color);

                    style.getTextNode(_fragmentsList[i].Text);

                    temp.InnerText = _fragmentsList[i].Text;
                    fontList.Add(temp);
                    //pNode.AppendChild(temp);
                    old_style = style;
                    continue;
                }

                if (_fragmentsList[i].StyleId != _fragmentsList[i - 1].StyleId)
                {
                    string BIU = style.getTextNode(_fragmentsList[i].Text);
                    if (old_style.FontColor != style.FontColor || !old_style.FontSize.Equals(style.FontSize) || !old_style.Font.Equals(style.Font))
                    {

                        if (!old_style.Font.Equals(style.Font))
                        {
                            XmlAttribute font = getXMLDocumentRoot().CreateAttribute("FACE");
                            font.Value = style.Font;
                            temp.Attributes.Append(font);
                        }
                        if (old_style.FontSize != style.FontSize)
                        {
                            XmlAttribute fontSize = getXMLDocumentRoot().CreateAttribute("SIZE");
                            fontSize.Value = style.FontSize.ToString();
                            temp.Attributes.Append(fontSize);
                        }
                        if (old_style.FontColor != style.FontColor)
                        {
                            XmlAttribute color = getXMLDocumentRoot().CreateAttribute("COLOR");
                            color.Value = style.FontColor.ToString();
                            temp.Attributes.Append(color);
                        }

                        temp.InnerText = BIU;
                        fontList.Add(temp);

                    }
                    else
                    {
                        fontList[i - 1].InnerText += BIU;
                    }
                }


                    old_style = style;
            }

            /*fontList.Reverse();

            XmlElement fontRoot = getXMLDocumentRoot().CreateElement("TEST");
            XmlElement oldItem = new XmlElement();

            int counter = 0;

            foreach (XmlElement item in fontList)
            {
                XmlElement temp = item;

                if (counter == 0)
                {
                    fontRoot = temp;
                    temp.AppendChild(oldItem);
                    oldItem = temp;
                    continue;
                }
                else
                {
                    temp.AppendChild(oldItem);
                    fontRoot = temp;
                    oldItem = temp;
                }
                
            }*/

            //pNode.AppendChild(fontRoot);

            return pNode;
        }

        public XmlElement getFragmentsNode()
        {
            XmlElement fragments = getXMLDocumentRoot().CreateElement("fragments");

            foreach (TextFragment tFragment in _fragmentsList)
            {
                XmlElement f = tFragment.getFragmentChild();
                fragments.AppendChild(f);
            }

            return fragments;
        }

        public XmlElement getStylesNode()
        {
            XmlElement styles = getXMLDocumentRoot().CreateElement("styles");

            foreach (TextStyle tStyle in _styleList)
            {
                XmlElement s = tStyle.getStylesChild();
                styles.AppendChild(s);
            }

            return styles;
        }

        public override XmlElement getXMLTree()
        {
            XmlElement parent = base.getXMLTree();
            XmlElement properties = getPropertiesNode();
            XmlElement acce = getXMLDocumentRoot().CreateElement("accessors");
            XmlElement textObjectProps = getTextObjectPropertiesNode();

            acce.AppendChild(textObjectProps);
            parent.AppendChild(properties);
            parent.AppendChild(acce);

            return parent;
        }

        public override XmlDocument getXMLDocumentRoot()
        {
            return base.getXMLDocumentRoot();
        }

        public override void setXMLDocumentRoot(ref XmlDocument xmldocument)
        {
            base.setXMLDocumentRoot(ref xmldocument);
        }

        public override Properties getProperties()
        {
            return base.getProperties();
        }

        public override void setProperties(Properties properties)
        {
            base.setProperties(properties);
        }

        public override void ConvertToYoobaUnits(int width, int height)
        {
            base.ConvertToYoobaUnits(width, height);

            //FONT, SIZE, COLOR, ALIGNMENT

            //Font size convertion
            _size /= 100;
            //_size = (int)Math.Round(_size * 1.5);

            //Font color
            _color = getFontColorAsInteger(_color).ToString();

            foreach(TextStyle style in StyleList)
            {
                //Font size convertion
                style.FontSize /= 100;
                //style.FontSize = (int) Math.Round(style.FontSize*1.5);

                //Fake font and color
                style.Font = "Arial";

            }

            //Alignment
            _align = (_align.ToLower() == "l" || _align.ToLower() == "left") ? "left" : _align;
            _align = (_align.ToLower() == "r" || _align.ToLower() == "right") ? "right" : _align;
            _align = (_align.ToLower() == "c" || _align.ToLower() == "ctr" || _align.ToLower() == "center") ? "center" : _align;
        
        }

        public int getFontColorAsInteger(string color)
        {
            if(color.Length == 6)
                return int.Parse(color, System.Globalization.NumberStyles.HexNumber);
            else
            {
                Console.WriteLine("Error in color convertion, '" + color + "' could not be converted!");
                return 0;
            }
        }

        public string Align
        {
            get { return _align; }
            set { _align = value; }
        }

        internal List<TextStyle> StyleList
        {
            get { return _styleList; }
            set { _styleList = value; }
        }

        internal List<TextFragment> FragmentsList
        {
            get { return _fragmentsList; }
            set { _fragmentsList = value; }
        }

        public string Autosize
        {
            get { return _autosize; }
            set { _autosize = value; }
        }

        public string Font
        {
            get { return _font; }
            set { _font = value; }
        }

        public string AntiAlias
        {
            get { return _antiAlias; }
            set { _antiAlias = value; }
        }

        public int Size
        {
            get { return _size; }
            set { _size = value; }
        }

        public int LetterSpacing
        {
            get { return _letterSpacing; }
            set { _letterSpacing = value; }
        }

        public int Leading
        {
            get { return _leading; }
            set { _leading = value; }
        }

        public string Color
        {
            get { return _color; }
            set { _color = value; }
        }

        public Boolean RunningText
        {
            get { return _runningText; }
            set { _runningText = value; }
        }

        public Boolean Selectable
        {
            get { return _selectable; }
            set { _selectable = value; }
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

        public Boolean UseScroller
        {
            get { return _useScroller; }
            set { _useScroller = value; }
        }


        internal void setAttributes(TableStyle tableStyle)
        {
            if (tableStyle == null)
                return;

            Color = (tableStyle.FontColor != "") ? tableStyle.FontColor : Color;
            Size = (tableStyle.FontSize != 0) ? tableStyle.FontSize : Size;
            Bold = tableStyle.Bold;
            Italic = tableStyle.Italic;
            Underline = tableStyle.Underline;
        }
    }
}
