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

        public ShapeObject(SceneObject sceneobject) : base(sceneobject) { }

        public override XmlElement getXMLTree()
        {
            Console.WriteLine("ShapeObject");
            return base.getXMLTree();
            
        }
        public override Properties getProperties()
        {
            return null;
        }
        
    }
}
