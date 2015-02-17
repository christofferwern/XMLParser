using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class ShapeObject : SceneObjectDecorator
    {
        private float _alpha, _fillAlpha, _gradientAngle, _rotation, _rotationX, _rotationY, _rotationZ;
        private int _lineAlpha, _lineSize, _x, _y, _z;

        private Boolean _cacheAsBitmap, _fillEnable, _lineEnable, _visible;
        private string _fillType, _gradientType, _fillColor, _lineColor;

        private float[] gradientAphas;
        private int[] gradientFills;

        private List<XmlElement> _accessorChildList;

        public ShapeObject(SceneObject sceneobject) : base(sceneobject) 
        {
            string objectType = "com.yooba.shapes.X";
            sceneobject.setObjectType(objectType);
            _accessorChildList = new List<XmlElement>();

            _alpha = 0;
            _cacheAsBitmap = false;
            _gradientAngle = 90;
            _gradientType = "linear";
            _fillAlpha = 0;
            _fillColor = "ffffff";
            _fillEnable = false;
            _fillType = "solid";
            _lineAlpha = 1;
            _lineColor = "000000";
            _lineEnable = true;
            _lineSize = 4;
            _rotation = 0;
            _rotationX = 0;
            _rotationY = 0;
            _rotationZ = 0;
            _visible = true;
            _x = 0;
            _y = 0;
            _z = 0;
        }

        private void generateAccessorChilds()
        {

            XmlElement alpha = getXMLDocumentRoot().CreateElement("alpha");
            alpha.InnerXml = _alpha.ToString();
            XmlElement cacheAsBitmap = getXMLDocumentRoot().CreateElement("cacheAsBitmap");
            cacheAsBitmap.InnerXml = _cacheAsBitmap.ToString();
            XmlElement gradientAngle = getXMLDocumentRoot().CreateElement("gradientAngle");
            gradientAngle.InnerXml = _gradientAngle.ToString();
            XmlElement gradientType = getXMLDocumentRoot().CreateElement("gradientType");
            gradientType.InnerXml = _gradientType.ToString();
            XmlElement fillAlpha = getXMLDocumentRoot().CreateElement("fillAlpha");
            fillAlpha.InnerXml = _fillAlpha.ToString();
            XmlElement fillColor = getXMLDocumentRoot().CreateElement("fillColor");
            fillColor.InnerXml = _fillColor.ToString();
            XmlElement fillType = getXMLDocumentRoot().CreateElement("fillType");
            fillType.InnerXml = _fillType.ToString();
            XmlElement fillEnable = getXMLDocumentRoot().CreateElement("fillEnable");
            fillEnable.InnerXml = _fillEnable.ToString();
            XmlElement lineAlpha = getXMLDocumentRoot().CreateElement("lineAlpha");
            lineAlpha.InnerXml = _lineAlpha.ToString();
            XmlElement lineColor = getXMLDocumentRoot().CreateElement("lineColor");
            lineColor.InnerXml = _lineColor.ToString();
            XmlElement lineEnable = getXMLDocumentRoot().CreateElement("lineEnable");
            lineEnable.InnerXml = _lineEnable.ToString();
            XmlElement lineSize = getXMLDocumentRoot().CreateElement("lineSize");
            lineSize.InnerXml = _lineSize.ToString();
            XmlElement rotation = getXMLDocumentRoot().CreateElement("rotation");
            rotation.InnerXml = _rotation.ToString();
            XmlElement rotationX = getXMLDocumentRoot().CreateElement("rotationX");
            rotationX.InnerXml = _rotationX.ToString();
            XmlElement rotationY = getXMLDocumentRoot().CreateElement("rotationY");
            rotationY.InnerXml = _rotationY.ToString();
            XmlElement rotationZ = getXMLDocumentRoot().CreateElement("rotationZ");
            rotationZ.InnerXml = _rotationZ.ToString();
            XmlElement visible = getXMLDocumentRoot().CreateElement("visible");
            visible.InnerXml = _visible.ToString();
            XmlElement x = getXMLDocumentRoot().CreateElement("x");
            x.InnerXml = _x.ToString();
            XmlElement y = getXMLDocumentRoot().CreateElement("y");
            y.InnerXml = _y.ToString();
            XmlElement z = getXMLDocumentRoot().CreateElement("z");
            z.InnerXml = _z.ToString();

            _accessorChildList.Add(alpha);
            _accessorChildList.Add(cacheAsBitmap);
            _accessorChildList.Add(gradientAngle);
            _accessorChildList.Add(gradientType);
            _accessorChildList.Add(fillAlpha);
            _accessorChildList.Add(fillColor);
            _accessorChildList.Add(fillType);
            _accessorChildList.Add(fillEnable);
            _accessorChildList.Add(lineAlpha);
            _accessorChildList.Add(lineColor);
            _accessorChildList.Add(lineEnable);
            _accessorChildList.Add(lineSize);
            _accessorChildList.Add(rotation);
            _accessorChildList.Add(rotationX);
            _accessorChildList.Add(rotationY);
            _accessorChildList.Add(rotationZ);
            _accessorChildList.Add(visible);
            _accessorChildList.Add(x);
            _accessorChildList.Add(y);
            _accessorChildList.Add(z);

        }

        public override XmlElement getXMLTree()
        {

            generateAccessorChilds();

            XmlElement parent = base.getXMLTree();
            XmlElement acce = getXMLDocumentRoot().CreateElement("accessors");
            foreach (XmlElement child in _accessorChildList)
            {
                acce.AppendChild(child);
            }

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

        public float RotationZ
        {
            get { return _rotationZ; }
            set { _rotationZ = value; }
        }

        public float RotationY
        {
            get { return _rotationY; }
            set { _rotationY = value; }
        }

        public float RotationX
        {
            get { return _rotationX; }
            set { _rotationX = value; }
        }

        public float Rotation
        {
            get { return _rotation; }
            set { _rotation = value; }
        }

        public float GradientAngle
        {
            get { return _gradientAngle; }
            set { _gradientAngle = value; }
        }
        public float FillAlpha
        {
            get { return _fillAlpha; }
            set { _fillAlpha = value; }
        }
        public float Alpha
        {
            get { return _alpha; }
            set { _alpha = value; }
        }

        public int Z
        {
            get { return _z; }
            set { _z = value; }
        }

        public int Y
        {
            get { return _y; }
            set { _y = value; }
        }

        public int X
        {
            get { return _x; }
            set { _x = value; }
        }

        public int LineSize
        {
            get { return _lineSize; }
            set { _lineSize = value; }
        }

        public string LineColor
        {
            get { return _lineColor; }
            set { _lineColor = value; }
        }

        public int LineAlpha
        {
            get { return _lineAlpha; }
            set { _lineAlpha = value; }
        }

        public string FillColor
        {
            get { return _fillColor; }
            set { _fillColor = value; }
        }

        public Boolean Visible
        {
            get { return _visible; }
            set { _visible = value; }
        }

        public Boolean LineEnable
        {
            get { return _lineEnable; }
            set { _lineEnable = value; }
        }

        public Boolean FillEnable
        {
            get { return _fillEnable; }
            set { _fillEnable = value; }
        }

        public Boolean CacheAsBitmap
        {
            get { return _cacheAsBitmap; }
            set { _cacheAsBitmap = value; }
        }

        public string GradientType
        {
            get { return _gradientType; }
            set { _gradientType = value; }
        }

        public string FillType
        {
            get { return _fillType; }
            set { _fillType = value; }
        }
        
    }
}
