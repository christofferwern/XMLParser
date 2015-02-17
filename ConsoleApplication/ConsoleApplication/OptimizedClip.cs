using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class OptimizedClip
    {
        private string _optID;
        private int _readSize, _writeSize, _optWidth, _optHeight;
        private bool _useOpt, _compress, _processed, _processFail;

        private string[] _optimizedClipAttributes = new string[9]{  "optID", "useOpt", "readSize", "writeSize", "compress", 
                                                                    "optWidth", "optHeight", "processed","processFail"};

        public OptimizedClip()
        {
            _optID = "";
            _useOpt = false;
            _readSize = 0;
            _writeSize = 0;
            _compress = true;
            _optWidth = 0;
            _optHeight = 0;
            _processed = false;
            _processFail = false;
        }

        public XmlElement getOptimizedClipNode(XmlDocument docRoot)
        {
            XmlElement optimizedClip = docRoot.CreateElement("optimizedClip");

            foreach (string s in _optimizedClipAttributes)
            {
                XmlAttribute xmlAttr = docRoot.CreateAttribute(s);

                FieldInfo fieldInfo = GetType().GetField("_" + s, BindingFlags.NonPublic | BindingFlags.Instance);
                if (fieldInfo != null)
                    xmlAttr.Value = fieldInfo.GetValue(this).ToString();

                optimizedClip.Attributes.Append(xmlAttr);
            }

            return optimizedClip;
        }
        public string OptID
        {
            get { return _optID; }
            set { _optID = value; }
        }
        public int OptHeight
        {
            get { return _optHeight; }
            set { _optHeight = value; }
        }

        public int OptWidth
        {
            get { return _optWidth; }
            set { _optWidth = value; }
        }

        public int WriteSize
        {
            get { return _writeSize; }
            set { _writeSize = value; }
        }
        public int ReadSize
        {
            get { return _readSize; }
            set { _readSize = value; }
        }

        public bool ProcessFail
        {
            get { return _processFail; }
            set { _processFail = value; }
        }

        public bool Processed
        {
            get { return _processed; }
            set { _processed = value; }
        }

        public bool Compress
        {
            get { return _compress; }
            set { _compress = value; }
        }

        public bool UseOpt
        {
            get { return _useOpt; }
            set { _useOpt = value; }
        }
    }
}
