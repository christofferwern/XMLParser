using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public abstract class SceneObjectDecorator : SceneObject
    {
        private SceneObject _sceneObject;

        public SceneObjectDecorator(SceneObject so)
        {
            _sceneObject = so;
        }

        public virtual XmlElement getXMLTree()
        {
            return _sceneObject.getXMLTree();
        }

        public virtual Properties getProperties()
        {
            return _sceneObject.getProperties();
        }

        public virtual void setProperties(Properties properties)
        {
            _sceneObject.setProperties(properties);
        }

        public virtual void setXMLDocumentRoot(ref XmlDocument xmldocument)
        {
            _sceneObject.setXMLDocumentRoot(ref xmldocument);
        }

        public virtual XmlDocument getXMLDocumentRoot()
        {
            return _sceneObject.getXMLDocumentRoot();
        }
    }
}
