using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class ImageObject : SceneObjectDecorator
    {

        public ImageObject(SceneObject sceneobject) : base(sceneobject) { }

        public override XmlElement getXMLTree()
        {
            XmlElement parent = base.getXMLTree();
            XmlElement acce = getXMLDocumentRoot().CreateElement("accessorsIMAGE");
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
    }
}
