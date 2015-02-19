using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public interface SceneObject
    {
        Properties getProperties();
        void setProperties(Properties p);
        XmlElement getXMLTree();
        void setXMLDocumentRoot(ref XmlDocument xmldocument);
        XmlDocument getXMLDocumentRoot();
        void setObjectType(string objectType);
        void ConvertToYoobaUnits(int width, int height);
    }
}
