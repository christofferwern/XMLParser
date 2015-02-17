using System;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class Properties
    {
        private Boolean _flip,
                        _formEnabled,
                        _fx,
                        _height,
                        _isActionExectuer,
                        _isCopyable,
                        _isMovable,
                        _isRemovable,
                        _isSwapable,
                        _opacity,
                        _rotation,
                        _text,
                        _width,
                        _xpos,
                        _ypos;


        public Properties()
        {
            _flip = false;
            _formEnabled = false;
            _fx = false;
            _height = false;
            _isActionExectuer = false;
            _isCopyable = false;
            _isMovable = false;
            _isRemovable = false;
            _isSwapable = false;
            _opacity = false;
            _rotation = false;
            _text = false;
            _width = false;
            _xpos = false;
            _ypos = false;
        }

        public Properties(  Boolean flip,
                            Boolean formEnabled,
                            Boolean fx,
                            Boolean height,
                            Boolean isActionExectuer,
                            Boolean isCopyable,
                            Boolean isMovable,
                            Boolean isRemovable,
                            Boolean isSwapable,
                            Boolean opacity,
                            Boolean rotation,
                            Boolean text,
                            Boolean width,
                            Boolean xpos,
                            Boolean ypos)
        {
            _flip = flip;
            _formEnabled = formEnabled;
            _fx = fx;
            _height = height;
            _isActionExectuer = isActionExectuer;
            _isCopyable = isCopyable;
            _isMovable = isMovable;
            _isRemovable = isRemovable;
            _isSwapable = isSwapable;
            _opacity = opacity;
            _rotation = rotation;
            _text = text;
            _width = width;
            _xpos = xpos;
            _ypos = ypos;
        }

        public Properties(SceneObject sceneObject)
        {
            Console.WriteLine();
        }

        public void setProperties(  Boolean flip,
                                    Boolean formEnabled,
                                    Boolean fx,
                                    Boolean height,
                                    Boolean isActionExectuer,
                                    Boolean isCopyable,
                                    Boolean isMovable,
                                    Boolean isRemovable,
                                    Boolean isSwapable,
                                    Boolean opacity,
                                    Boolean rotation,
                                    Boolean text,
                                    Boolean width,
                                    Boolean xpos,
                                    Boolean ypos)
        {
            _flip = flip;
            _formEnabled = formEnabled;
            _fx = fx;
            _height = height;
            _isActionExectuer = isActionExectuer;
            _isCopyable = isCopyable;
            _isMovable = isMovable;
            _isRemovable = isRemovable;
            _isSwapable = isSwapable;
            _opacity = opacity;
            _rotation = rotation;
            _text = text;
            _width = width;
            _xpos = xpos;
            _ypos = ypos;
        }

        public Boolean Ypos
        {
            get { return _ypos; }
            set { _ypos = value; }
        }

        public Boolean Xpos
        {
            get { return _xpos; }
            set { _xpos = value; }
        }

        public Boolean Width
        {
            get { return _width; }
            set { _width = value; }
        }

        public Boolean Text
        {
            get { return _text; }
            set { _text = value; }
        }

        public Boolean Rotation
        {
            get { return _rotation; }
            set { _rotation = value; }
        }

        public Boolean Opacity
        {
            get { return _opacity; }
            set { _opacity = value; }
        }

        public Boolean IsSwapable
        {
            get { return _isSwapable; }
            set { _isSwapable = value; }
        }

        public Boolean IsRemovable
        {
            get { return _isRemovable; }
            set { _isRemovable = value; }
        }

        public Boolean IsMovable
        {
            get { return _isMovable; }
            set { _isMovable = value; }
        }

        public Boolean IsCopyable
        {
            get { return _isCopyable; }
            set { _isCopyable = value; }
        }

        public Boolean IsActionExectuer
        {
            get { return _isActionExectuer; }
            set { _isActionExectuer = value; }
        }

        public Boolean Height
        {
            get { return _height; }
            set { _height = value; }
        }

        public Boolean Fx
        {
            get { return _fx; }
            set { _fx = value; }
        }

        public Boolean FormEnabled
        {
            get { return _formEnabled; }
            set { _formEnabled = value; }
        }

        public Boolean Flip
        {
            get { return _flip; }
            set { _flip = value; }
        }

        public XmlElement getNode(XmlDocument root)
        {
            XmlElement node = root.CreateElement("properties");
            node.InnerText = toString();

            return node;
        }
        public string toString()
        {
            string output = "";

            output += (_flip) ? "flip," : "";
            output += (_formEnabled) ? "formEnabled," : "";
            output += (_fx) ? "fx," : "";
            output += (_height) ? "height," : "";
            output += (_isActionExectuer) ? "isActionExectuer," : "";
            output += (_isCopyable) ? "isCopyable," : "";
            output += (_isMovable) ? "isMovable," : "";
            output += (_isRemovable) ? "isRemovable," : "";
            output += (_isSwapable) ? "isSwapable," : "";
            output += (_opacity) ? "opacity," : "";
            output += (_rotation) ? "rotation," : "";
            output += (_text) ? "text," : "";
            output += (_width) ? "width," : "";
            output += (_xpos) ? "xpos," : "";
            output += (_ypos) ? "ypos," : "";

            if (output.Contains(","))
                output = output.Remove(output.LastIndexOf(","));

            return output;
        }



    }
}
