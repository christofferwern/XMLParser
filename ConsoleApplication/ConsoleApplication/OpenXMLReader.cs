using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;

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

        private const int _nrOfThemeColors = 16;
        private int _presentationSizeX, _presentationSizeY;
        private string[,] _themeColors;
        private List<DrawingML.Outline> _themeLines;
        private List<PowerPointText> _slideMasterPowerPointShapes;
        private PresentationDocument _presentationDocument;

        private XmlDocument _rootXmlDoc;

        internal PresentationObject PresentationObject
        {
            get { return _presentationObject; }
            set { _presentationObject = value; }
        }

        public OpenXMLReader(string path)
        {
            _path = path;
            _presentationObject = new PresentationObject();

            _themeColors = new string[_nrOfThemeColors, 2];
            _themeLines = new List<DrawingML.Outline>();
            _slideMasterPowerPointShapes = new List<PowerPointText>();

            _presentationObject.getXmlDocument(out _rootXmlDoc);

            _presentationSizeX = 0;
            _presentationSizeY = 0;

        }

        public string Path
        {
            get { return _path; }
            set { _path = value; }
        }

        public void read()
        {
           // try
           // {
                // Open the presentation file as read-only.
                using (_presentationDocument = PresentationDocument.Open(_path, false))
                {
                    //Retrive the presentation part
                    var presentation = _presentationDocument.PresentationPart.Presentation;

                    //Get the size of presentation
                    PresentationML.SlideSize slideInfo = presentation.SlideSize;
                    _presentationSizeX = slideInfo.Cx.Value;
                    _presentationSizeY = slideInfo.Cy.Value;


                    //Get the slidemaster scene object, background etc
                    PresentationML.SlideMaster slideMaster = _presentationDocument.PresentationPart.SlideMasterParts.ElementAt(0).SlideMaster;;

                    foreach (var child in slideMaster.CommonSlideData.ChildElements)
                    {
                        if (child.LocalName == "bg")
                            _presentationObject.BackgroundSceneObjectList.AddRange(getSceneObjects((PresentationML.Background)child));

                        if (child.LocalName == "spTree")
                            _presentationObject.BackgroundSceneObjectList.AddRange(getSceneObjects((PresentationML.ShapeTree)child));
                    }


                    //Counter of scenes
                    int sceneCounter = 1;

                    //Go through all Slides in the PowerPoint presentation
                    foreach (PresentationML.SlideId slideID in presentation.SlideIdList)
                    {
                        SlidePart slidePart = _presentationDocument.PresentationPart.GetPartById(slideID.RelationshipId) as SlidePart;
                        Scene scene = new Scene(sceneCounter);

                        //Go through all elements in the slide layout
                        foreach (var child in slidePart.SlideLayoutPart.SlideLayout.CommonSlideData.ChildElements)
                        {
                            if (child.LocalName == "bg")
                                scene.addSceneObjects(getSceneObjects((PresentationML.Background)child));

                            if (child.LocalName == "spTree")
                                scene.addSceneObjects(getSceneObjects((PresentationML.ShapeTree)child));
                        }

                        //Go through all elements in the slide
                        foreach(var child in slidePart.Slide.CommonSlideData.ChildElements)
                        {
                            if(child.LocalName == "bg")
                                scene.addSceneObjects(getSceneObjects((PresentationML.Background)child));
                            
                            if(child.LocalName == "spTree")
                                scene.addSceneObjects(getSceneObjects((PresentationML.ShapeTree)child));
                        }

                        //Add scene to presentation
                        _presentationObject.addScene(scene);

                        sceneCounter++;
                    }
                }
           // }
           // catch
           // {
           //     Console.WriteLine("Error reading: " + _path);
           // }
        }

        private List<SceneObject> getSceneObjects(PresentationML.ShapeTree shapeTree)
        {
            List<SceneObject> sceneObjectList = new List<SceneObject>();

            foreach (var child in shapeTree.ChildElements)
            {
                if (child.LocalName == "sp")
                    sceneObjectList.AddRange(getSceneObjects((PresentationML.Shape)child));
                

                if (child.LocalName == "graphicFrame")
                    sceneObjectList.AddRange(getSceneObjects((PresentationML.GraphicFrame)child));
                

                if (child.LocalName == "pic")
                    sceneObjectList.AddRange(getSceneObjects((PresentationML.Picture)child));
                

                if (child.LocalName == "grpSp")
                    sceneObjectList.AddRange(getSceneObjects((PresentationML.GroupShape)child));
                
            }

            return sceneObjectList;
        }

        private List<SceneObject> getSceneObjects(PresentationML.GroupShape groupShape)
        {
            List<SceneObject> sceneObjectList = new List<SceneObject>();
            return sceneObjectList;
        }

        private List<SceneObject> getSceneObjects(PresentationML.Picture picture)
        {
            List<SceneObject> sceneObjectList = new List<SceneObject>();
            return sceneObjectList;
        }

        private List<SceneObject> getSceneObjects(PresentationML.GraphicFrame graphicFrame)
        {
            List<SceneObject> sceneObjectList = new List<SceneObject>();
            return sceneObjectList;
        }

        private List<SceneObject> getSceneObjects(PresentationML.Shape shape)
        {
            List<SceneObject> sceneObjectList = new List<SceneObject>();
            return sceneObjectList;
        }

        private List<SceneObject> getSceneObjects(PresentationML.Background background)
        {
            List<SceneObject> sceneObjectList = new List<SceneObject>();
            return sceneObjectList;
        }


        private List<string> getColors(DrawingML.GradientFill gradFill)
        {
            List<string> colors = new List<string>();

            return colors;
        }

        private string getColor(DrawingML.SolidFill solidFill)
        {
            string color = "";

            return color;
        }

    }
}
