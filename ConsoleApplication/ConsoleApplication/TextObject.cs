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

        public TextObject(SceneObject sceneobject) : base(sceneobject) { }

        public override XmlElement getXMLTree()
        {
            
            XmlElement XE =  base.getXMLTree();
            
            return XE;
        }

        public string Align
        {
            get { return _align; }
            set { _align = value; }
        }

    }
}
