using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class ShapeObject : SceneObjectDecorator
    {
        private float _alpha, _cornerRadius, _fillAlpha, _fillAlpha1, _fillAlpha2, _gradientAngle, _rotation, _rotationX, _rotationY, _rotationZ, _scaleZ;

        private int _lineAlpha, _lineSize, _x, _y, _z, _points, _radius;

        private Boolean _cacheAsBitmap, _fillEnable, _lineEnable, _visible;
        private string _fillType, _gradientType, _fillColor, _lineColor, _fillColor1, _fillColor2;

        private float[] _gradientAlphas;
        private string[] _gradientFills;

        private string[] _attributes = new string[24]{   "alpha", "cacheAsBitmap", "cornerRadius", "gradientAngle", "gradientType", "fillAlpha", "fillColor",
                                                         "fillEnable", "fillType", "lineAlpha", "lineColor", "lineEnable", "lineSize", "points", "radius",
                                                         "rotation", "rotationX", "rotationY", "rotationZ", "scaleZ", "visible", "x",
                                                         "y", "z"};

        private List<String> _shapeObjectAccessorChild = new List<string>();
        private List<XmlElement> _accessorChildList;
        private Properties _properties;

        public enum shape_type{
            Rectangle,
            Circle,
            Polygon
        }

        public ShapeObject(SceneObject sceneobject, shape_type shapeType) : base(sceneobject) 
        {
            _shapeObjectAccessorChild.AddRange(_attributes);
            
            string objectType = "";

            switch (shapeType)
            {
                case shape_type.Rectangle:
                    objectType = "com.yooba.shapes.RoundedRectangleShape";
                    _shapeObjectAccessorChild.Remove("radius");
                    //Console.WriteLine("Rect");
                break;
                case shape_type.Circle:
                    objectType = "com.yooba.shapes.CircleShape";
                    _shapeObjectAccessorChild.Remove("cornerRadius");
                    //Console.WriteLine("Rect");
                break;
                case shape_type.Polygon:
                    objectType = "com.yooba.shapes.PolygonShape";
                    _shapeObjectAccessorChild.Remove("radius");
                    _shapeObjectAccessorChild.Remove("cornerRadius");
                    //Console.WriteLine("Rect");
                break;

            }

            sceneobject.setObjectType(objectType);
            _accessorChildList = new List<XmlElement>();

            _alpha = 1;
            _cacheAsBitmap = false;
            _cornerRadius = 0;
            _gradientAngle = 90;
            _gradientType = "linear";
            _fillAlpha = 0;
            _fillColor = "16743690";
            _fillEnable = true;
            _fillType = "solid";
            _lineAlpha = 1;
            _lineColor = "6426397";
            _lineEnable = false;
            _lineSize = 1;
            _points = 3;
            _radius = 0;
            _rotation = 0;
            _rotationX = 0;
            _rotationY = 0;
            _rotationZ = 0;
            _scaleZ = 1;
            _visible = true;
            _x = 0;
            _y = 0;
            _z = 0;
            _gradientAlphas = new float[2];
            _gradientFills = new string[2];

            _fillColor1 = _fillColor;
            _fillColor2 = "10834182";
            _fillAlpha1 = 1;
            _fillAlpha2 = 1;

            _properties = new Properties(true,false,true,true,true,true,true,true,false,true,true,false,true,true,true);
        }

        public override XmlElement getXMLTree()
        {

            XmlElement parent = base.getXMLTree();

            parent.AppendChild(Properties.getNode(getXMLDocumentRoot()));

            XmlElement acce = getXMLDocumentRoot().CreateElement("accessors");

            foreach (string s in _shapeObjectAccessorChild)
            {
                XmlElement xmlChild = getXMLDocumentRoot().CreateElement(s);

                FieldInfo fieldInfo = GetType().GetField("_" + s, BindingFlags.NonPublic | BindingFlags.Instance);
                if (fieldInfo != null)
                    xmlChild.InnerText = fieldInfo.GetValue(this).ToString().Replace(",",".").ToLower();

                Console.WriteLine(xmlChild.InnerText);

                acce.AppendChild(xmlChild);
            }

            XmlElement gradientColor = getXMLDocumentRoot().CreateElement("gradientFills");
            gradientColor.InnerText = _fillColor1.ToString() + ", " + _fillColor2.ToString();
            XmlElement gradientAlpha = getXMLDocumentRoot().CreateElement("gradientAlphas");
            gradientAlpha.InnerText = _fillAlpha1.ToString().Replace(",", ".") + ", " + _fillAlpha2.ToString().Replace(",", ".");

            acce.AppendChild(gradientColor);
            acce.AppendChild(gradientAlpha);

            parent.AppendChild(acce);

            return parent;
            
        }

        public override void ConvertToYoobaUnits(int width, int height)
        {
            base.ConvertToYoobaUnits(width, height);
            if (_fillType.Equals("solid"))
            {
                Console.WriteLine("S: " + _fillAlpha);
                _fillColor = getColorAsInteger(_fillColor).ToString();
                Console.WriteLine("I: " + _fillAlpha);
                _fillAlpha = 1 - (_fillAlpha/10000);
                Console.WriteLine("A: " + _fillAlpha);
            }

            if (_fillType.Equals("gradient"))
            {
                _fillColor = getColorAsInteger(_gradientFills[0]).ToString();

                if (_gradientAlphas[0] != 1)
                    _fillAlpha1 = ((_gradientAlphas[0] / 1000) / 100);
                if (_gradientAlphas[1] != 1)
                    _fillAlpha2 = ((_gradientAlphas[1] / 1000) / 100);

                _fillColor1 = _fillColor;
                _fillColor2 = getColorAsInteger(_gradientFills[1]).ToString();

                _gradientAngle /= 60000; 
            }
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

        public int getColorAsInteger(string color)
        {
            return int.Parse(color, System.Globalization.NumberStyles.HexNumber);
        }

        public float CornerRadius
        {
            get { return _cornerRadius; }
            set { _cornerRadius = value; }
        }

        public float ScaleZ
        {
            get { return _scaleZ; }
            set { _scaleZ = value; }
        }

        public string FillColor2
        {
            get { return _fillColor2; }
            set { _fillColor2 = value; }
        }

        public string FillColor1
        {
            get { return _fillColor1; }
            set { _fillColor1 = value; }
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

        public float[] GradientAlphas
        {
            get { return _gradientAlphas; }
            set { _gradientAlphas = value; }
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

        public string[] GradientFills
        {
            get { return _gradientFills; }
            set { _gradientFills = value; }
        }

        public int Radius
        {
            get { return _radius; }
            set { _radius = value; }
        }

        public int Points
        {
            get { return _points; }
            set { _points = value; }
        }

        public Properties Properties
        {
            get { return _properties; }
            set { _properties = value; }
        }
    }
}
