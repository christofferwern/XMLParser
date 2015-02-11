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

        //BoundsX and BoundsY corresponds to positions from left top corner
        //ClipHeight and ClipWidth are the height and width of the scene object
        private int _boundsX, _boundsY, _clipHeight, _clipWidth;

        private string _clipID, _name, _type;

        private Boolean _hidden;

        private XmlDocument doc;

        private Properties _properties;

        public SimpleSceneObject()
        {
            _clipID = Guid.NewGuid().ToString();
            _name = "SimpleSceneObject";
            doc = new XmlDocument();
            _properties = new Properties();
            _alpha = 1;
            _hidden = false;
            _width = 0;
            _height = 0;
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
            doc.AppendChild(XE);
            return XE;
        }

        public float Z
        {
            get { return _z; }
            set { _z = value; }
        }

        public float Rotation
        {
            get { return _rotation; }
            set { _rotation = value; }
        }

        public float Width
        {
            get { return _width; }
            set { _width = value; }
        }

        public float Height
        {
            get { return _height; }
            set { _height = value; }
        }

        public float Alpha
        {
            get { return _alpha; }
            set { _alpha = value; }
        }

        public int ClipWidth
        {
            get { return _clipWidth; }
            set { _clipWidth = value; }
        }

        public int ClipHeight
        {
            get { return _clipHeight; }
            set { _clipHeight = value; }
        }

        public int BoundsY
        {
            get { return _boundsY; }
            set { _boundsY = value; }
        }

        public int BoundsX
        {
            get { return _boundsX; }
            set { _boundsX = value; }
        }

        public string Type
        {
            get { return _type; }
            set { _type = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public string ClipID
        {
            get { return _clipID; }
            set { _clipID = value; }
        }

        public Boolean Hidden
        {
            get { return _hidden; }
            set { _hidden = value; }
        }

    }
}
