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
        private float _alpha, _cornerRadius, _gradientAngle, _rotation, _rotationX, _rotationY, _rotationZ;
        private int _fillAlpha, _fillColor, _lineAlpha, _lineColor, _lineSize, _x, _y;
        private Boolean _cacheAsBitmap, _fillEnable, _lineEnable, _visible;
        private string _fillType, _gradientType;

        private float[] gradientAphas;
        private int[] gradientFills;

        public ShapeObject(SceneObject sceneobject) : base(sceneobject) 
        {
            string objectType = "com.yoob.shapes.X";
            sceneobject.setObjectType(objectType);

        }

        public override XmlElement getXMLTree()
        {
            XmlElement parent = base.getXMLTree();
            XmlElement acce = getXMLDocumentRoot().CreateElement("accessors");
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
        
    }
}
