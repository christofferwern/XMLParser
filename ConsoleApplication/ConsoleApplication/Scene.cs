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
        public  Properties Properties
        {
            get { return _properties; }
            set { _properties = value; }
        }

        private string _sceneLabel;
        public string SceneLabel
        {
            get { return _sceneLabel; }
            set { _sceneLabel = value; }
        }
        //private List<SceneObject> sceneObjectList;

        public Scene()
        {
            _properties = new Properties();
            _sceneLabel = "Scene Label";
        //    sceneObjectList = new List<SceneObject>();
        }

        //public void addSceneObject(SceneObject sceneObject)
        //{
        //    sceneObjectList.Add(sceneObject);
        //}

        //public void removeSceneObject(SceneObject sceneObject)
        //{
        //    sceneObjectList.Remove(sceneObject);
        //}

        public void getXMLTree()
        {

        }

    }
}
