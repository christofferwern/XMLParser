using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class ImageResponse
    {
        string _clipID, _itemType, _width, _height, _label, _size, _uploadDate, _url;


        public ImageResponse(){
            _clipID = "";
            _itemType = "";
            _width = "";
            _height = "";
            _label = "";
            _size = "";
            _uploadDate = "";
            _url = "";
        }

        public string Url
        {
            get { return _url; }
            set { _url = value; }
        }

        public string UploadDate
        {
            get { return _uploadDate; }
            set { _uploadDate = value; }
        }

        public string Size
        {
            get { return _size; }
            set { _size = value; }
        }

        public string Label
        {
            get { return _label; }
            set { _label = value; }
        }

        public string Height
        {
            get { return _height; }
            set { _height = value; }
        }

        public string Width
        {
            get { return _width; }
            set { _width = value; }
        }

        public string ItemType
        {
            get { return _itemType; }
            set { _itemType = value; }
        }

        public string ClipID
        {
            get { return _clipID; }
            set { _clipID = value; }
        }

        public string toString()
        {
            return "clipID: " +_clipID +
                     "\nitemType : " + _itemType +
                     "\nwidth : " + _width +
                     "\nheight : " + _height +
                     "\nlabel : " + _label +
                     "\nsize : " + _size +
                     "\nuploadDate : " + _uploadDate +
                     "\nurl : " + _url;
        }
    }
}
