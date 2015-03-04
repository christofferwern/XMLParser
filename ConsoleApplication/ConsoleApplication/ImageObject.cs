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

        /*<sceneObject type="FluddClip" clipID="4938d5b0-0a24-4977-b218-f7612f6e4630" z="0" boundsX="209" boundsY="183" clipWidth="200" clipHeight="200" width="0" height="0" rotation="0" alpha="1" name="31D3B859-8325-3FDA-F85E-CD5C1B97FAD2" hidden="false" flip="0">
            <optimizedClip optID="" useOpt="false" readSize="0" writeSize="0" compress="true" optWidth="0" optHeight="0" processed="false" processFail="false"/>
            <dsCol><![CDATA[]]></dsCol>
            <properties>opacity,fx,xpos,ypos,width,height,rotation,flip,isRemovable,isMovable,isCopyable,isActionExecuter,isSwappable</properties>
          </sceneObject>*/

        private Properties _properties;

        public ImageObject(SceneObject sceneobject) : base(sceneobject) 
        {
            sceneobject.setObjectType("FluddClip");
            _properties = new Properties(true, false, true, true, true, true, true, true, true, true, true, false, true, true, true);
        }

        public XmlElement getOptimizedClip()
        {
            return null;
        }

        public override XmlElement getXMLTree()
        {
            XmlElement parent = base.getXMLTree();
            XmlElement properties = _properties.getNode(getXMLDocumentRoot());
            parent.AppendChild(properties);

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
