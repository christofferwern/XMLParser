using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class SimpleSceneObject : SceneObject
    {
        private float _alpha, _height, _width, _rotation, _z;
        private int _boundsX, _boundsY, _clipHeight, _clipWidth;
        private string _clipID, _name, _type;
        private Boolean _hidden;
        private XmlDocument doc;

        public SimpleSceneObject()
        {
            _name = "SimpleScenObject";
            doc = new XmlDocument();
        }

        public float Alpha
        {
            get { return _alpha; }
            set { _alpha = value; }
        }

        /*public Propterties getProperties()
        {
            throw new NotImplementedException();
        }

        public void setProperties(Properties p)
        {
            throw new NotImplementedException();
        }*/

        public XmlElement getXMLTree()
        {
            XmlElement XE = doc.DocumentElement;
            return XE;
        }
    }
}
