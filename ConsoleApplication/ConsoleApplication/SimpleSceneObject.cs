﻿using System;
using System.Reflection;
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
        private int _boundsX, _boundsY, _clipHeight, _clipWidth, _flip;
        private string _clipID, _type, _name;
        private Boolean _hidden;
        private XmlDocument _doc;
        private Properties _properties;
        private OptimizedClip _optimizedClip;

        private string[] _attributes = new string[14]{  "type", "clipID","z", "boundsX", "boundsY", "clipWidth", "clipHeight",
                                                        "width","height", "rotation", "alpha","name", "hidden", "flip"};
        private SimpleSceneObject simpleSceneObjectShape;

        public SimpleSceneObject()
        {
            _clipID = Guid.NewGuid().ToString();
            _name = Guid.NewGuid().ToString();
            _doc = new XmlDocument();
            _alpha = 1;
            _hidden = false;
            _width = 0;
            _height = 0;
            _z = 0;
            _flip = 0;
            _properties = new Properties();
            _optimizedClip = new OptimizedClip();

            _rotation = 0;
            _z = 0;
            _type = "";
            _boundsX = 0;
            _boundsY = 0;
            _clipHeight = 0;
            _clipWidth = 0;
        }

        public SimpleSceneObject(SimpleSceneObject s)
        {
            _clipID = s.ClipID;
            _name = s.Name;
            _doc = new XmlDocument();
            _alpha = s.Alpha;
            _hidden = s.Hidden;
            _width = s.Width;
            _height = s.Height;
            _z = s.Z;
            _flip = 0;
            _properties = new Properties();
            _optimizedClip = new OptimizedClip();

            _rotation = s.Rotation;
            _z = s.Z;
            _type = s.Type;
            _boundsX = s.BoundsX;
            _boundsY = s.BoundsY;
            _clipHeight = s.ClipHeight;
            _clipWidth = s.ClipWidth;
        }

        public object Clone()
        {
            return this;
        }

        public void ConvertToYoobaUnits(int width, int height)
        {
            //_boundsX, _boundsY, _clipHeight, _clipWidth,
            int pptWidth = width, pptHeight = height, yoobaWidth = 1024, yoobaHeight = 768;

            float scaleWidth = (float)yoobaWidth / (float)pptWidth, scaleHeight = (float)yoobaHeight / (float)pptHeight;

            _boundsX = (int) Math.Round(_boundsX * scaleWidth);
            _boundsY = (int) Math.Round(_boundsY * scaleHeight);
            _clipWidth = (int) Math.Round(_clipWidth * scaleWidth);
            _clipHeight = (int) Math.Round(_clipHeight * scaleHeight);

            //rotation
            _rotation = _rotation / 60000;

            //Handle negative and too large angles
            while (_rotation < 0)
                _rotation += 360;

            while (_rotation > 360)
                _rotation -= 360;

            double tempRot = 0;

            //Handle the 4 different cases
            if (_rotation >= 0 && _rotation <= 90)
                tempRot = _rotation;
            else if (_rotation > 90 && _rotation <= 180)
                tempRot = 180 - _rotation;
            else if (_rotation > 180 && _rotation <= 270)
                tempRot = _rotation - 180;
            else if (_rotation > 270 && _rotation < 360)
                tempRot = 360 - _rotation;
            
            //Store the Center of mass for the object before the rotation
            double COM_X = _boundsX + _clipWidth / 2, 
                   COM_Y = _boundsY + _clipHeight / 2;

            //Calculate the top left position for the rotated object, stored in newX and newY
            double newAngle = 90 - tempRot;
            double newAngleInRadians = newAngle * Math.PI / 180;
            double stepLeft = Math.Cos(newAngleInRadians) * ClipHeight;
            double newX = stepLeft + BoundsX;
            double newY = BoundsY;

            //Calculate the center of mass position for the rotated object
            newX += Math.Sin(newAngleInRadians) * ClipWidth / 2;
            newY += Math.Cos(newAngleInRadians) * ClipWidth / 2;
            newX += Math.Sin(newAngleInRadians - Math.PI / 2) * ClipHeight / 2;
            newY += Math.Cos(newAngleInRadians - Math.PI / 2) * ClipHeight / 2;

            //Calculate the difference in COM
            double diffX = newX - COM_X;
            double diffY = newY - COM_Y;

            //Subtract the diffrence from the original bounds
            _boundsX -= (int)Math.Round(diffX);
            _boundsY -= (int)Math.Round(diffY);

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
        }

        public void setObjectType(string objectType)
        {
            _type = objectType;
        }

        public XmlDocument getXMLDocumentRoot()
        {
            return _doc;
        }

        public XmlElement getXMLTree()
        {

            //generateAttributes();
            XmlElement xmlElement = _doc.CreateElement("sceneObject");

            xmlElement.AppendChild(_optimizedClip.getOptimizedClipNode(_doc));
            xmlElement.AppendChild(getDsColNode());

            foreach (string s in _attributes)
            {
                XmlAttribute xmlAttr = _doc.CreateAttribute(s);

                FieldInfo fieldInfo = GetType().GetField("_" + s, BindingFlags.NonPublic | BindingFlags.Instance);
                if (fieldInfo != null)
                    xmlAttr.Value = fieldInfo.GetValue(this).ToString();

                xmlElement.Attributes.Append(xmlAttr);
            }

            _doc.DocumentElement.AppendChild(xmlElement);

            return xmlElement;
        }

        public XmlElement getDsColNode()
        {
            XmlElement dsCol = getXMLDocumentRoot().CreateElement("dsCol");
            XmlCDataSection cData = getXMLDocumentRoot().CreateCDataSection("");

            dsCol.AppendChild(cData);

            return dsCol;
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
