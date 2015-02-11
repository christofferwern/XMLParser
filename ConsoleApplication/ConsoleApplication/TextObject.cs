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

        public TextObject(SceneObject sceneobject) : base(sceneobject) { }

        public override XmlElement getXMLTree()
        {
            XmlElement parent = base.getXMLTree();
            XmlElement acce = getXMLDocumentRoot().CreateElement("accessorsTEXT");
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

    }
}
