﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class PowerPointText
    {
        private int _fontSize, _x, _y, _cx, _cy, _rotation, _idx;
        private string _type, _font, _alignment, _anchor, _fontColor;
        private Boolean _bold, _italic, _underline;

        public PowerPointText()
        {
            _idx = -1;
            _fontSize = 0;
            _x = 0;
            _y = 0;
            _cx = 0;
            _cy = 0;
            _rotation = 0;
            _fontColor = "";
            _type = "";
            _font = "";
            _alignment = "";
            _anchor = "";
            _bold = false;
            _italic = false;
            _underline = false;
        }

        public PowerPointText(PowerPointText ppt)
        {
            _idx = ppt.Idx;
            _fontSize = ppt.FontSize;
            _x = ppt.X;
            _y = ppt.Y;
            _cx = ppt.Cx;
            _cy = ppt.Cy;
            _rotation = ppt.Rotation;
            _fontColor = ppt.FontColor;
            _type = ppt.Type;
            _font = ppt.Font;
            _alignment = ppt.Alignment;
            _anchor = ppt.Anchor;
            _bold = ppt.Bold;
            _italic = ppt.Italic;
            _underline = ppt.Underline; 
        }

        public String toString()
        {
            return "Placeholder: (" + _type + ", " + _idx + ")\n" +
                   "  Font:        " + _font + "\n" +
                   "  Font size:   " + _fontSize + "\n" +
                   "  Font color:  " + _fontColor + "\n" +
                   "  Size:        (" + _cx + "," + _cy + ")\n" +
                   "  Position:    (" + _x + "," + _y + ")\n" +
                   "  Rotation:    " + Rotation + "\n" +
                   "  Anchor:      " + _anchor + "\n" +
                   "  Alignment:   " + _alignment + "\n" + 
                   "  B U I:       (" + _bold + ", " + _underline + ", " + _italic + ") \n";
        }

        public int Idx
        {
            get { return _idx; }
            set { _idx = value; }
        }

        public Boolean Underline
        {
            get { return _underline; }
            set { _underline = value; }
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

        public string Anchor
        {
            get { return _anchor; }
            set { _anchor = value; }
        }

        public string Alignment
        {
            get { return _alignment; }
            set { _alignment = value; }
        }

        public string Font
        {
            get { return _font; }
            set { _font = value; }
        }

        public string Type
        {
            get { return _type; }
            set { _type = value; }
        }
        
        public int Rotation
        {
            get { return _rotation; }
            set { _rotation = value; }
        }

        public int Cy
        {
            get { return _cy; }
            set { _cy = value; }
        }

        public int Cx
        {
            get { return _cx; }
            set { _cx = value; }
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

        public int FontSize
        {
            get { return _fontSize; }
            set { _fontSize = value; }
        }

        public string FontColor
        {
            get { return _fontColor; }
            set { _fontColor = value; }
        }

        public bool isEmpty()
        {
            if( _idx == -1 &&
                _fontSize == 0 &&
                //_x == 0 &&
                //_y == 0 &&
                //_cx == 0 &&
                //_cy == 0 &&
                _rotation == 0 &&
                _fontColor == "" &&
                _type == "" &&
                _font == "" &&
                //_alignment == "" &&
                //_anchor == "" &&
                _bold == false &&
                _italic == false &&
                _underline == false
            )
                return true;
            else
                return false;
        }
    }
}
