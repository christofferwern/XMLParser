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

    }
}
