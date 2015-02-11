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

        private XmlDocument _doc;

        private Properties _properties;

        private List<XmlAttribute> _objectAttributes;

        public SimpleSceneObject()
        {
            _clipID = Guid.NewGuid().ToString();
            _name = "SimpleSceneObject";
            _doc = new XmlDocument();
            _alpha = 1;
            _hidden = false;
            _width = 0;
            _height = 0;
            _properties = new Properties();
            _objectAttributes = new List<XmlAttribute>();
        }

        private void generateAttributes()
        {
            XmlAttribute type = _doc.CreateAttribute("type");
            type.Value = "";
            XmlAttribute clipID = _doc.CreateAttribute("clipID");
            clipID.Value = "";
            XmlAttribute z = _doc.CreateAttribute("z");
            z.Value = "";
            XmlAttribute boundsX = _doc.CreateAttribute("boundsX");
            boundsX.Value = "";
            XmlAttribute boundsY = _doc.CreateAttribute("boundsY");
            boundsY.Value = "";
            XmlAttribute clipWidth = _doc.CreateAttribute("clipWidth");
            clipWidth.Value = "";
            XmlAttribute clipHeight = _doc.CreateAttribute("clipHeight");
            clipHeight.Value = "";
            XmlAttribute width = _doc.CreateAttribute("width");
            width.Value = "";
            XmlAttribute height = _doc.CreateAttribute("height");
            height.Value = "";
            XmlAttribute rotation = _doc.CreateAttribute("rotation");
            rotation.Value = "";
            XmlAttribute alpha = _doc.CreateAttribute("alpha");
            alpha.Value = "";
            XmlAttribute name = _doc.CreateAttribute("name");
            name.Value = "";
            XmlAttribute hidden = _doc.CreateAttribute("hidden");
            hidden.Value = "";
            XmlAttribute flip = _doc.CreateAttribute("flip");
            flip.Value = "";

            _objectAttributes.Add(type);
            _objectAttributes.Add(clipID);
            _objectAttributes.Add(z);
            _objectAttributes.Add(boundsX);
            _objectAttributes.Add(boundsY);
            _objectAttributes.Add(clipWidth);
            _objectAttributes.Add(clipHeight);
            _objectAttributes.Add(width);
            _objectAttributes.Add(height);
            _objectAttributes.Add(rotation);
            _objectAttributes.Add(alpha);
            _objectAttributes.Add(name);
            _objectAttributes.Add(hidden);
            _objectAttributes.Add(flip);
        }

        public Properties getProperties()
        {
            return _properties;
        }

        public void setProperties(Properties properties)
        {
            _properties = properties;
        }

        public void setXMLDocumentRoot(ref XmlDocument xmldocument)
        {
            _doc = xmldocument;
            //generateAttributes();
        }

        public XmlDocument getXMLDocumentRoot()
        {
            return _doc;
        }

        public XmlElement getXMLTree()
        {
            XmlElement XE = _doc.CreateElement("sceneObject");

            foreach (XmlAttribute XA in _objectAttributes)
                XE.Attributes.Append(XA);
            

            _doc.DocumentElement.AppendChild(XE);
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
