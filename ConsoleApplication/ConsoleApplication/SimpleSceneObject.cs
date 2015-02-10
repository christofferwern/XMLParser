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

        private Properties _properties;

        public SimpleSceneObject()
        {
            _name = "SimpleScenObject";
            doc = new XmlDocument();
            _properties = new Properties();
        }

        public float Alpha
        {
            get { return _alpha; }
            set { _alpha = value; }
        }

        public Properties getProperties()
        {
            return _properties;
        }

        public void setProperties(Properties properties)
        {
            _properties = properties;
        }

        public XmlElement getXMLTree()
        {
            XmlElement XE = doc.CreateElement("root");
            return XE;
        }
    }
}
