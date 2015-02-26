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

                    _presentationObject.ConvertToYoobaUnits(_presentationSizeX, _presentationSizeY);
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
            SimpleSceneObject simpleSceneObject = new SimpleSceneObject();
            ShapeObject shapeObject = new ShapeObject(simpleSceneObject, ShapeObject.shape_type.Rectangle);


            //**TODO**
            //Get siblings and check for offsets!!!


            bool HasBg = false, HasLine = false, HasValidGeometry = false;
            foreach (var child in shape.ChildElements)
            {
                if(child.LocalName == "nvSpPr")
                {

                }

                if(child.LocalName == "spPr")
                {
                    PresentationML.ShapeProperties spPr = (PresentationML.ShapeProperties)child;

                    if (spPr.Transform2D != null)
                    {
                        simpleSceneObject.BoundsX = (spPr.Transform2D.Offset.X != null) ? (int)spPr.Transform2D.Offset.X : simpleSceneObject.BoundsX;
                        simpleSceneObject.BoundsY = (spPr.Transform2D.Offset.Y != null) ? (int)spPr.Transform2D.Offset.Y : simpleSceneObject.BoundsY;
                        simpleSceneObject.ClipWidth = (spPr.Transform2D.Extents.Cx != null) ? (int)spPr.Transform2D.Extents.Cx : simpleSceneObject.ClipWidth;
                        simpleSceneObject.ClipHeight = (spPr.Transform2D.Extents.Cy != null) ? (int)spPr.Transform2D.Extents.Cy : simpleSceneObject.ClipHeight;
                    }

                    foreach(var spPrChild in child)
                    {
                        if(spPrChild.LocalName == "prstGeom")
                        {
                            DrawingML.PresetGeometry prstGeom = (DrawingML.PresetGeometry)spPrChild;

                            if (prstGeom.Preset == "rect"           || prstGeom.Preset == "snip1Rect"       || prstGeom.Preset == "snip2DiagRect"   ||
                                prstGeom.Preset == "round1Rect"     || prstGeom.Preset == "round2DiagRect"  || prstGeom.Preset == "round2SameRect"  ||
                                prstGeom.Preset == "snip2SameRect"  || prstGeom.Preset == "snipRoundRect"   || prstGeom.Preset == "round1Rect"      ||
                                prstGeom.Preset == "flowChartProcess")
                            {
                                HasValidGeometry = true;
                                shapeObject = new ShapeObject(simpleSceneObject, ShapeObject.shape_type.Rectangle);
                            }


                            if (prstGeom.Preset == "ellipse" || prstGeom.Preset == "flowChartConnector")
                            {
                                shapeObject = new ShapeObject(simpleSceneObject, ShapeObject.shape_type.Circle);
                                HasValidGeometry = true;
                            }

                            if (prstGeom.Preset == "triangle" || prstGeom.Preset == "diamond" || prstGeom.Preset == "flowChartDecision" || prstGeom.Preset == "flowChartExtract" ||
                                prstGeom.Preset == "pentagon" || prstGeom.Preset == "hexagon" || prstGeom.Preset == "heptagon" || prstGeom.Preset == "flowChartPreparation" ||
                                prstGeom.Preset == "octagon" || prstGeom.Preset == "decagon" || prstGeom.Preset == "dodecagon" || prstGeom.Preset == "flowChartMerge")
                            {
                                HasValidGeometry = true;

                                if (prstGeom.Preset == "flowChartMerge")
                                {
                                    simpleSceneObject.Rotation = 180 * 60000;
                                    shapeObject.Rotation += 180 * 60000;
                                }

                                shapeObject = new ShapeObject(simpleSceneObject, ShapeObject.shape_type.Polygon);
                                shapeObject.Points = (prstGeom.Preset == "triangle" || prstGeom.Preset == "flowChartExtract" || prstGeom.Preset == "flowChartMerge") ? 3 :
                                                     (prstGeom.Preset == "diamond" || prstGeom.Preset == "flowChartDecision") ? 4 :
                                                     (prstGeom.Preset == "pentagon") ? 5 :
                                                     (prstGeom.Preset == "hexagon" || prstGeom.Preset == "flowChartPreparation") ? 6 :
                                                     (prstGeom.Preset == "heptagon") ? 7 :
                                                     (prstGeom.Preset == "octagon") ? 8 :
                                                     (prstGeom.Preset == "decagon") ? 10 :
                                                     12;
                            }


                        }

                        if(spPrChild.LocalName == "solidFill")
                        {
                            HasBg = true;
                            shapeObject.FillColor = getColor((DrawingML.SolidFill)spPrChild);
                        }   

                        if(spPrChild.LocalName == "gradFill")
                        {
                            HasBg = true;
                            List<string> colors = getColors((DrawingML.GradientFill)spPrChild);
                            shapeObject.FillColor1 = colors[0];
                            shapeObject.FillColor2 = colors[colors.Count-1];
                        }  

                        if(spPrChild.LocalName == "noFill")
                        {
                        
                        } 

                        if(spPrChild.LocalName == "blipFill")
                        {
                        
                        }  

                        if(spPrChild.LocalName == "ln")
                        {


                            HasLine = true;
                        }
                    }
                }

                if(child.LocalName == "style")
                {
                    PresentationML.ShapeStyle style = (PresentationML.ShapeStyle)child;
                    
                    foreach(var styleChild in child)
                    {
                        if(styleChild.LocalName == "lnRef")
                        {
                            DrawingML.LineReference lnRef = (DrawingML.LineReference)styleChild;
                        }

                        if(styleChild.LocalName == "fillRef")
                        {
                            DrawingML.FillReference fillRef = (DrawingML.FillReference)styleChild;
                        }

                        if(styleChild.LocalName == "fontRef")
                        {
                            DrawingML.FontReference fontRef = (DrawingML.FontReference)styleChild;
                        }
                    }

                }

                if(child.LocalName == "txBody")
                {

                }
            }


            if (HasValidGeometry && (HasLine || HasBg))
                sceneObjectList.Add(shapeObject);

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

        private string getColor(OpenXmlElement openXmlElement)
        {
            string color = "";

            return color;
        }

    }
}
