using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class PresentationObject
    {
        private int _backgroundColor, _projektWidth, _projectHeight;
        private string _id;
        private List<Scene> sceneList;
        private XmlDocument _xmlDoc;
        List<SceneObject> _backgroundSceneObjectList;

        public PresentationObject()
        {
            _projektWidth = 1024;
            _projectHeight = 768;

            _id = Guid.NewGuid().ToString();
            _backgroundColor = Color.Red.ToArgb();
            _backgroundSceneObjectList = new List<SceneObject>();
            sceneList = new List<Scene>();
            _xmlDoc = new XmlDocument();
        }

        public void addScene(Scene scene)
        {
            sceneList.Add(scene);
        }

        public void removeScene(Scene scene)
        {
            sceneList.Remove(scene);
        }

        public void getXmlDocument(out XmlDocument outs)
        {
            outs = _xmlDoc;
        }

        public void createXMLTree()
        {
            
            

        }

        public XmlElement getBackgroundSceneNode()
        {
            XmlElement background = _xmlDoc.CreateElement("backgroundScene");
            XmlElement properties = _xmlDoc.CreateElement("properties");
            XmlElement backgroundColor = _xmlDoc.CreateElement("bgColor");
            backgroundColor.InnerXml = "15298";
            XmlAttribute bgMode = _xmlDoc.CreateAttribute("backgroundMode");
            bgMode.Value = "";

            background.Attributes.Append(bgMode);
            background.AppendChild(properties);
            background.AppendChild(backgroundColor);


            int z_index = 0;
            foreach (SceneObject item in _backgroundSceneObjectList)
            {
                item.setZindex(z_index);
                item.setXMLDocumentRoot(ref _xmlDoc);
                background.AppendChild(item.getXMLTree());
                z_index++;
            }

            return background;
        }

        public XmlElement getForegroundSceneNode()
        {
            XmlElement foreground = _xmlDoc.CreateElement("foregroundScene");
            XmlElement properties = _xmlDoc.CreateElement("properties");

            foreground.AppendChild(properties);

            return foreground;
        }
        public XmlElement getSceneTransitionNode()
        {
            XmlElement sceneTransition = _xmlDoc.CreateElement("sceneTransition");
            XmlAttribute swf = _xmlDoc.CreateAttribute("swf");
            swf.Value = "CrossFade";

            XmlElement transSettings = _xmlDoc.CreateElement("transSettings");
            XmlAttribute type = _xmlDoc.CreateAttribute("type");
            type.Value = "YoobaTransition";
            XmlAttribute trans_swf = _xmlDoc.CreateAttribute("swf");
            trans_swf.Value = "trans.YoobaTransition";

            XmlElement duration = _xmlDoc.CreateElement("duration");
            duration.InnerXml = "0.6";
            XmlElement direction = _xmlDoc.CreateElement("direction");
            direction.InnerXml = "0";
            XmlElement direction_2 = _xmlDoc.CreateElement("direction");
            direction_2.InnerXml = "0";
            XmlElement color = _xmlDoc.CreateElement("color");
            color.InnerXml = "16777215";
            XmlElement motionBlur = _xmlDoc.CreateElement("motionBlur");
            motionBlur.InnerXml = "true";

            transSettings.Attributes.Append(type);
            transSettings.Attributes.Append(trans_swf);

            transSettings.AppendChild(duration);
            transSettings.AppendChild(direction);
            transSettings.AppendChild(direction_2);
            transSettings.AppendChild(color);
            transSettings.AppendChild(motionBlur);

            sceneTransition.Attributes.Append(swf);

            sceneTransition.AppendChild(transSettings);

            return sceneTransition;
        }

        public XmlDocument getXMLTree()
        {

            XmlDeclaration xmlDeclaration = _xmlDoc.CreateXmlDeclaration("1.0", "utf-8", null);
            XmlElement root = _xmlDoc.CreateElement("yoobaProject");
            _xmlDoc.InsertBefore(xmlDeclaration, _xmlDoc.DocumentElement);
            _xmlDoc.AppendChild(root);

            root.AppendChild(getSceneTransitionNode());
            root.AppendChild(getBackgroundSceneNode());
            root.AppendChild(getForegroundSceneNode());

            foreach(Scene scene in sceneList){
                scene.setXMLDocumentRoot(ref _xmlDoc);
                _xmlDoc.DocumentElement.AppendChild(scene.getXMLTree());
            }
            return _xmlDoc;
        }

        public List<SceneObject> BackgroundSceneObjectList
        {
            get { return _backgroundSceneObjectList; }
            set { _backgroundSceneObjectList = value; }
        }

        public int ProjectHeight
        {
            get { return _projectHeight; }
            set { _projectHeight = value; }
        }

        public int ProjektWidth
        {
            get { return _projektWidth; }
            set { _projektWidth = value; }
        }

        public string Id
        {
            get { return _id; }
            set { _id = value; }
        }

        public int BackgroundColor
        {
            get { return _backgroundColor; }
            set { _backgroundColor = value; }
        }

        public void ConvertToYoobaUnits(int width, int height)
        {
            foreach(Scene scene in this.sceneList)
                foreach (SceneObject sceneObject in scene.SceneObjectList)
                    sceneObject.ConvertToYoobaUnits(width, height);

            foreach (SceneObject sceneObject in BackgroundSceneObjectList)
                sceneObject.ConvertToYoobaUnits(width, height);

        }
    }
}
