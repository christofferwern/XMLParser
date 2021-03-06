﻿using System;
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

        private Boolean _cacheAsBitmap, _fillEnable, _lineEnabled, _visible;
        private string _fillType, _gradientType, _fillColor, _lineColor, _fillColor1, _fillColor2;

        private float[] _gradientAlphas;
        private string[] _gradientFills;

        private string[] _attributes = new string[24]{   "alpha", "cacheAsBitmap", "cornerRadius", "gradientAngle", "gradientType", "fillAlpha", "fillColor",
                                                         "fillEnable", "fillType", "lineAlpha", "lineColor", "lineEnabled", "lineSize", "points", "radius",
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
                break;
                case shape_type.Circle:
                    objectType = "com.yooba.shapes.CircleShape";
                    _shapeObjectAccessorChild.Remove("cornerRadius");
                break;
                case shape_type.Polygon:
                    objectType = "com.yooba.shapes.PolygonShape";
                    _shapeObjectAccessorChild.Remove("radius");
                    _shapeObjectAccessorChild.Remove("cornerRadius");
                break;

            }

            sceneobject.setObjectType(objectType);
            _accessorChildList = new List<XmlElement>();

            _alpha = 1;
            _cacheAsBitmap = false;
            _cornerRadius = 0;
            _gradientAngle = 90;
            _gradientType = "linear";
            _fillAlpha = 100000;
            _fillAlpha1 = 100000;
            _fillAlpha2 = 100000;
            _fillColor = "16743690";
            _fillEnable = true;
            _fillType = "solid";
            _lineAlpha = 1;
            _lineColor = "6426397";
            _lineEnabled = false;
            _lineSize = 0;
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
            _gradientAlphas[0] = 100000;
            _gradientAlphas[1] = 100000;
            _gradientFills = new string[2];

            _fillColor1 = _fillColor;
            _fillColor2 = "10834182";

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
                _fillColor = getColorAsInteger(_fillColor).ToString();
                _fillAlpha /= 100000;
            }

            if (_fillType.Equals("gradient"))
            {

                _fillColor = getColorAsInteger(_fillColor1).ToString();

                _fillAlpha1 /= 100000;
                _fillAlpha2 /= 100000;

                _fillColor1 = _fillColor;
                _fillColor2 = getColorAsInteger(_fillColor2).ToString();
                
                _gradientAngle /= 60000;

            }

            _lineSize = (int) Math.Round((double)_lineSize / 12700);

            if (_lineSize <= 0)
                _lineEnabled = false;
            else
                _lineEnabled = true;

            _lineColor = getColorAsInteger(_lineColor).ToString();
            _cornerRadius = (float) Math.Round((_cornerRadius / 100000) * 128 * 4);

            _rotation /= 60000;
            
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
            if (color != "")
                return int.Parse(color, System.Globalization.NumberStyles.HexNumber);
            else
                return 0;
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

        public Boolean LineEnabled
        {
            get { return _lineEnabled; }
            set { _lineEnabled = value; }
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

        public float FillAlpha2
        {
            get { return _fillAlpha2; }
            set { _fillAlpha2 = value; }
        }

        public float FillAlpha1
        {
            get { return _fillAlpha1; }
            set { _fillAlpha1 = value; }
        }

        public Properties Properties
        {
            get { return _properties; }
            set { _properties = value; }
        }

        internal void setAttributes(TableStyle tableStyle)
        {
            if(tableStyle==null)
                return;

            FillAlpha = (tableStyle.FillAlpha != 0) ? tableStyle.FillAlpha : FillAlpha;
            FillColor = (tableStyle.FillColor != "") ? tableStyle.FillColor : FillColor;
            LineSize = (tableStyle.LineSize != 0) ? tableStyle.LineSize : LineSize;
            LineColor = (tableStyle.LineColor != "") ? tableStyle.LineColor : LineColor;
        }
    }
}
