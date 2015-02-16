using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class TextObject : SceneObjectDecorator
    {
        private string _align, _antiAlias, _font, _autosize;
        private List<TextStyle> _styleList;
        private List<TextFragment> _fragmentsList;
        private int _color, _leading, _letterSpacing, _size;
        private Boolean _bold, _italic, _underLine, _selectable, _runningText;
        private XmlDocument _doc;

        public TextObject(SceneObject sceneobject) : base(sceneobject) 
        {
            string objectType = "com.customObjects.TextObject";
            _styleList = new List<TextStyle>();
            _fragmentsList = new List<TextFragment>();

            setProperties(new Properties(true, false, true, true, true, true, true, true,
                                         false, true, true, true, true, true, true));

            sceneobject.setObjectType(objectType);
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
            XmlElement stylesNode = getStylesNode();
            XmlElement fragmentNode = getFragmentsNode();
            XmlElement textNode = getTextNode();

            textObjectPropNode.AppendChild(textNode);
            /*textObjectPropNode.AppendChild(stylesNode);
            textObjectPropNode.AppendChild(fragmentNode);*/

            return textObjectPropNode;
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
            XmlElement acce = getXMLDocumentRoot().CreateElement("accessors");
            XmlElement textObjectProps = getTextObjectPropertiesNode();

            acce.AppendChild(textObjectProps);
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

    }
}
