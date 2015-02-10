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
        /*Propterties getProperties();
        void setProperties(Properties p);
        */
        XmlElement getXMLTree();
    }
}
