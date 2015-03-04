using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace ConsoleApplication
{
    class MediaFile
    {
        private int _offX, _offY, _extX, _extY, _rotation, _alpha;


        private string _imageType, _imageLocation, _storeLocation, _imageName;
        private ImagePart _imagePart;

        public MediaFile()
        {
            _extX = 0;
            _extY = 0;
            _offX = 0;
            _offY = 0;
            _alpha = 100000;
            _rotation = 0;
            _imageType = "";
            _imagePart = null;
            _storeLocation = @"C:\Users\ex1\Desktop\PPTImages\";
        }

        public string ImageName
        {
            get { return _imageName; }
            set 
            { 
                _imageName = value;
                ImageType = _imageName.Split('.')[1]; 
            }
        }

        public string StoreLocation
        {
            get { return _storeLocation; }
            set { _storeLocation = value; }
        }

        public string ImageLocation
        {
            get { return _imageLocation; }
            set { _imageLocation = value; }
        }

        public string ImageType
        {
            get { return _imageType; }
            set { _imageType = value; }
        }

        public int Alpha
        {
            get { return _alpha; }
            set { _alpha = value; }
        }

        public int Rotation
        {
            get { return _rotation; }
            set { _rotation = value; }
        }

        public int ExtY
        {
            get { return _extY; }
            set { _extY = value; }
        }

        public int ExtX
        {
            get { return _extX; }
            set { _extX = value; }
        }

        public int OffY
        {
            get { return _offY; }
            set { _offY = value; }
        }

        public int OffX
        {
            get { return _offX; }
            set { _offX = value; }
        }

        public ImagePart ImagePart
        {
            get { return _imagePart; }
            set { _imagePart = value; }
        }

        public string toString()
        {
            return "Image information \n" +
                   " Image Name: " + _imageName + "\n" +
                   " Offset Position (X,Y): (" + _offX + "," + _offY + ")\n" +
                   " Extent Position (X,Y): (" + _extX + "," + _extY + ")\n" +
                   " Rotation: " + _rotation + "\n" +
                   " Image Type: " + _imageType + "\n" +
                   " Alpha: " + _alpha + "\n";
        }


    }
}
