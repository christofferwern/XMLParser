using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class Scene
    {
        private Properties _properties;
        private List<SceneObject> _sceneObjectList;
        private string _sceneLabel;
        private XmlDocument _doc;



        public  Properties Properties
        {
            get { return _properties; }
            set { _properties = value; }
        }
        
        public string SceneLabel
        {
            get { return _sceneLabel; }
            set { _sceneLabel = value; }
        }


        public Scene(int sceneNumber)
        {
            _properties = new Properties();
            _sceneLabel = "Scene " + sceneNumber.ToString();
            _sceneObjectList = new List<SceneObject>();

            //_sceneObjectList.Add(new TextObject(new SimpleSceneObject()));
            //_sceneObjectList.Add(new ShapeObject(new SimpleSceneObject()));
            //_sceneObjectList.Add(new TextObject(new ShapeObject(new SimpleSceneObject())));
            
        }

        public void addSceneObject(SceneObject sceneObject)
        {
            //HÄR NÅGONSTANS VILL VI SORTERA EFTER Z-INDEX
            _sceneObjectList.Add(sceneObject);
        }

        public void addSceneObjects(List<SceneObject> list)
        {
            foreach(SceneObject sceneObject in list)
                _sceneObjectList.Add(sceneObject);
        }

        public void removeSceneObject(SceneObject sceneObject)
        {
            _sceneObjectList.Remove(sceneObject);
        }

        public XmlElement getXMLTree()
        {
            XmlElement _rootElement = _doc.CreateElement("scene");
            XmlAttribute sceneLabel = _doc.CreateAttribute("sceneLabel");
            sceneLabel.Value = _sceneLabel;
            _rootElement.Attributes.Append(sceneLabel);
            _doc.DocumentElement.AppendChild(_rootElement);

            foreach (SceneObject scenObject in _sceneObjectList)
            {
                scenObject.setXMLDocumentRoot(ref _doc);
                _rootElement.AppendChild(scenObject.getXMLTree());

            }
           
            return _rootElement;
        }

        public void setXMLDocumentRoot(ref XmlDocument xmlDocument)
        {
            _doc = xmlDocument;
        }

    }
}
