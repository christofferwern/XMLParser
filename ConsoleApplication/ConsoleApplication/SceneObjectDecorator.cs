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

        /*public Propterties getProperties()
        {
            throw new NotImplementedException();
        }

        public void setProperties(Properties p)
        {
            throw new NotImplementedException();
        }*/

        public virtual XmlElement getXMLTree()
        {
            return _sceneObject.getXMLTree();
        }
    }
}
