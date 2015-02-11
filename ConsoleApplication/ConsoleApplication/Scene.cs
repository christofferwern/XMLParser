using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class Scene
    {
        private Properties _properties;
        private List<SceneObject> _sceneObjectList;
        private string _sceneLabel;

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


        public Scene()
        {
            _properties = new Properties();
            _sceneLabel = "Scene Label";
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

        public void getXMLTree()
        {

        }

    }
}
