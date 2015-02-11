﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class TextObject : SceneObjectDecorator
    {
        private string _align, _antiAlias, _font, _autosize;
        private List<TextStyle> _styleList;
        private List<TextFragment> _fragmentsList;
        private int _color, _leading, _letterSpacing, _size;
        private Boolean _bold, _italic, _underLine, _selectable, _runningText;

        public TextObject(SceneObject sceneobject) : base(sceneobject) 
        { 
            _styleList = new List<TextStyle>();
            _fragmentsList = new List<TextFragment>();
        }

        public override XmlElement getXMLTree()
        {
            
            XmlElement XE =  base.getXMLTree();
            
            return XE;
        }

        public string Align
        {
            get { return _align; }
            set { _align = value; }
        }

        public string Autosize
        {
            get { return _autosize; }
            set { _autosize = value; }
        }

        public string Font
        {
            get { return _font; }
            set { _font = value; }
        }

        public string AntiAlias
        {
            get { return _antiAlias; }
            set { _antiAlias = value; }
        }

        public int Size
        {
            get { return _size; }
            set { _size = value; }
        }

        public int LetterSpacing
        {
            get { return _letterSpacing; }
            set { _letterSpacing = value; }
        }

        public int Leading
        {
            get { return _leading; }
            set { _leading = value; }
        }

        public int Color
        {
            get { return _color; }
            set { _color = value; }
        }

        public Boolean RunningText
        {
            get { return _runningText; }
            set { _runningText = value; }
        }

        public Boolean Selectable
        {
            get { return _selectable; }
            set { _selectable = value; }
        }

        public Boolean UnderLine
        {
            get { return _underLine; }
            set { _underLine = value; }
        }

        public Boolean Italic
        {
            get { return _italic; }
            set { _italic = value; }
        }

        public Boolean Bold
        {
            get { return _bold; }
            set { _bold = value; }
        }

        internal List<TextStyle> StyleList
        {
            get { return _styleList; }
            set { _styleList = value; }
        }

        internal List<TextFragment> FragmentsList
        {
            get { return _fragmentsList; }
            set { _fragmentsList = value; }
        }
    }
}
