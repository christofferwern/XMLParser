using System;
using System.Reflection;
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
        private string _sceneLabel, _name, _dataSourceID, _sceneType, _numNavIndex;
        private bool _sceneIsScrollable, _enableTimeTracking;
        private int _preloadType, _dataSourceRecord;

        private string[] _scenAttributes = new string[] { "sceneLabel", "sceneIsScrollable", "numNavIndex", "dataSourceID", "dataSourceRecord",
                                                          "name", "sceneType", "enableTimeTracking", "preloadType"};

        private XmlDocument _doc;

        public Scene(int sceneNumber)
        {
            /********ATTRIBUTES***********/
            if (sceneNumber == 1)
                _name = "SceneViewer1855";
            else
                _name = Guid.NewGuid().ToString();

            _dataSourceID = "";
            _dataSourceRecord = 1;
            _sceneType = "blank";
            _sceneLabel = "Scene " + sceneNumber.ToString();
            _sceneIsScrollable = false;
            _enableTimeTracking = false;
            _preloadType = 0;
            _numNavIndex = "";
            /******END OF ATTRIBUTES******/

            _properties = new Properties();
            _properties.IsRemovable = true;
            _properties.IsMovable = true;
            _properties.FormEnabled = true;

            _sceneObjectList = new List<SceneObject>();
            
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

            foreach (string s in _scenAttributes)
            {
                XmlAttribute xmlAttr = _doc.CreateAttribute(s);

                FieldInfo fieldInfo = GetType().GetField("_" + s, BindingFlags.NonPublic | BindingFlags.Instance);
                if (fieldInfo != null)
                    xmlAttr.Value = fieldInfo.GetValue(this).ToString();

                _rootElement.Attributes.Append(xmlAttr);
            }


            _rootElement.AppendChild(_properties.getNode(_doc));

            foreach (SceneObject sceneObject in _sceneObjectList)
            {
                sceneObject.setXMLDocumentRoot(ref _doc);
                _rootElement.AppendChild(sceneObject.getXMLTree());

            }

            //_rootElement.AppendChild(getFormNode());

            return _rootElement;
        }

        private XmlElement getFormNode()
        {
            XmlElement form = _doc.CreateElement("form");

            form.Attributes.Append(_doc.CreateAttribute("id"));
            form.Attributes.Append(_doc.CreateAttribute("name"));
            form.Attributes.Append(_doc.CreateAttribute("description"));

            return form;
        }

        public void setXMLDocumentRoot(ref XmlDocument xmlDocument)
        {
            _doc = xmlDocument;
        }

        public List<SceneObject> SceneObjectList
        {
            get { return _sceneObjectList; }
            set { _sceneObjectList = value; }
        }
        public Properties Properties
        {
            get { return _properties; }
            set { _properties = value; }
        }

        public string SceneLabel
        {
            get { return _sceneLabel; }
            set { _sceneLabel = value; }
        }


    }
}
