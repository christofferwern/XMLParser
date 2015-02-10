using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class PresentationObject
    {
        private int _backgroundColor, _projektWidth, _projectHeight;
        private string _id;
        private List<Scene> sceneList;
        List<SceneObject> backgroundSceneObjectList;

        public PresentationObject()
        {
            _projektWidth = 1024;
            _projectHeight = 768;

            _id = Guid.NewGuid().ToString();
            _backgroundColor = Color.Red.ToArgb();
            sceneList = new List<Scene>();
        }

        public void addScene(Scene scene)
        {
            sceneList.Add(scene);
        }

        public void removeScene(Scene scene)
        {
            sceneList.Remove(scene);
        }

        public void createXMLTree()
        {

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
    }
}
