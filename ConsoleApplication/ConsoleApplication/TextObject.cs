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
        private string _align, _antiAlias, _font, _autosize;
        private List<TextStyle> _styleList;
        private List<TextFragment> _fragmentsList;
        private int _color, _leading, _letterSpacing, _size;
        private Boolean _bold, _italic, _underline, _selectable, _runningText, _useScroller;
        private XmlDocument _doc;
        private string[] _textObjectPropertiesAttributes = new string[14]{   "font", "align", "color", "italic", "bold", "underline", "size", "runningText",
                                                                             "autosize", "leading", "letterSpacing", "antiAlias", "useScroller", "selectable"};

        public TextObject(SceneObject sceneobject) : base(sceneobject) 
        {
            _align = "";
            _antiAlias = "";
            _font = "";
            _autosize = "";
            _color = 0;
            _leading = 0;
            _letterSpacing = 0;
            _size = 0;
            _bold = false;
            _italic = false;
            _underline = false;
            _selectable = false;
            _runningText = false;
            _useScroller = false;

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

            _doc = base.getXMLDocumentRoot();

            foreach (string s in _textObjectPropertiesAttributes)
            {
                XmlAttribute xmlAttr = _doc.CreateAttribute(s);

                FieldInfo fieldInfo = GetType().GetField("_" + s, BindingFlags.NonPublic | BindingFlags.Instance);
                if (fieldInfo!=null)
                    xmlAttr.Value = fieldInfo.GetValue(this).ToString();

                textObjectPropNode.Attributes.Append(xmlAttr);
            }

            textObjectPropNode.AppendChild(stylesNode);
            textObjectPropNode.AppendChild(fragmentNode);

            return textObjectPropNode;
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

        public int Color
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
    }
}
