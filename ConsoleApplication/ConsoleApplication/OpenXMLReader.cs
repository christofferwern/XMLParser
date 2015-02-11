using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.Serialization.Json;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using PresentationML = DocumentFormat.OpenXml.Presentation;
using DrawingML = DocumentFormat.OpenXml.Drawing;

namespace ConsoleApplication
{
    class OpenXMLReader
    {
        private string _path;
        private PresentationObject _presentationObject;

        public OpenXMLReader(string path)
        {
            _path = path;
            _presentationObject = new PresentationObject();
        }

        public string Path
        {
            get { return _path; }
            set { _path = value; }
        }

        public void read()
        {
            // Open the presentation file as read-only.
            using (PresentationDocument presentationDocument = PresentationDocument.Open(_path, false))
            {
                //Retrive the presentation part
                PresentationPart presentationPart = presentationDocument.PresentationPart;

                //Go through all Slides in the PowerPoint presentation
                foreach (SlidePart slidePart in presentationPart.SlideParts)
                {
                    Scene scene = new Scene();


                    //Get all text shape
                    List<SceneObject> textShapelist = GetTextShapesFromSlidePart(slidePart);

                    //Get all images
                    List<SceneObject> imageShapelist = GetImageShapesFromSlidePart(slidePart);

                    //Get background
                    List<SceneObject> backgroundShapelist = GetBackgroundShapesFromSlidePart(slidePart);

                    //Add all scene object to the scene
                    scene.addSceneObjects(textShapelist);
                    scene.addSceneObjects(imageShapelist);
                    scene.addSceneObjects(backgroundShapelist);

                    //Add scene to presentation
                    _presentationObject.addScene(scene);
                }
            }
        }

        private List<SceneObject> GetBackgroundShapesFromSlidePart(SlidePart slidePart)
        {
            List<SceneObject> backgroundShapelist = new List<SceneObject>();

            return backgroundShapelist;
        }

        private List<SceneObject> GetImageShapesFromSlidePart(SlidePart slidePart)
        {
            List<SceneObject> imageShapelist = new List<SceneObject>();

            return imageShapelist;
        }

        private List<SceneObject> GetTextShapesFromSlidePart(SlidePart slidePart)
        {
            List<SceneObject> textShapelist = new List<SceneObject>();

            return textShapelist;
        }
    }
}
