using System;
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
            _presentationObject.getXmlDocument(out _rootXmlDoc);
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
                foreach (SlidePart slidePart in presentationPart.SlideParts.Reverse())
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
            List<SceneObject> sceneObjectList = new List<SceneObject>();

            //Get the shape tree, <p:spTree>, which contains all shapes in the slide
            var spTree = slidePart.Slide.CommonSlideData.ShapeTree;

            //Traverse through all the shapes in the tree
            foreach (PresentationML.Shape sp in spTree.Descendants<PresentationML.Shape>())
            {
                SimpleSceneObject simpleSceneObject = new SimpleSceneObject();

                //Get info about SimpleSceneObject

                //Get the rotation of scene object
                foreach (var xfrm in sp.Elements<DrawingML.Transform2D>())
                {
                    simpleSceneObject.Rotation = (xfrm.Rotation != null) ? xfrm.Rotation : simpleSceneObject.Rotation;
                }

                //Get the positions of the scene object
                foreach (var off in sp.Descendants<DrawingML.Offset>())
                {
                    simpleSceneObject.BoundsX = (off.X != null) ? (int)off.X : simpleSceneObject.BoundsX;
                    simpleSceneObject.BoundsY = (off.Y != null) ? (int)off.Y : simpleSceneObject.BoundsY;
                }

                //Get the size of the shape object
                foreach (var ext in sp.Descendants<DrawingML.Extents>())
                {
                    simpleSceneObject.ClipHeight = (ext.Cx != null) ? (int)ext.Cx : simpleSceneObject.ClipHeight;
                    simpleSceneObject.ClipWidth = (ext.Cy != null) ? (int)ext.Cy : simpleSceneObject.ClipWidth;
                }

                //Add text info to scene object
                TextObject sceneObject = new TextObject(simpleSceneObject);

                //Go trough all the runs in the text body
                //they contains the text and some properties
                foreach (DrawingML.Run run in sp.TextBody.Descendants<DrawingML.Run>())
                {
                    //If there is a shape without any text, continue to next shape
                    if (run.Text.Text.Equals(""))
                        continue;

                    TextFragment textFragment = new TextFragment();
                    textFragment.setXMLDocumentRoot(ref _rootXmlDoc);
                    textFragment.Text = run.Text.Text;
                    
                    TextStyle textStyle = new TextStyle();
                    textStyle.setXMLDocumentRoot(ref _rootXmlDoc);

                    //Get font
                    foreach (var symbolFont in run.Elements<DrawingML.SymbolFont>())
                    {
                        textStyle.Font = symbolFont.Typeface;
                    }

                    //Get the texy body color, if it has been changed manually in the ppt file.
                    foreach (var color in run.Elements<DrawingML.RgbColorModelHex>())
                    {
                        //Convert Hexadeciamal color to integer color
                        textStyle.FontColor = int.Parse(color.Val, System.Globalization.NumberStyles.HexNumber);
                    }

                    //Get run properties (size, bold, italic, underline) and insert into style
                    textStyle.FontSize  = (run.RunProperties.FontSize != null)  ? (int)run.RunProperties.FontSize   : textStyle.FontSize;
                    textStyle.Bold      = (run.RunProperties.Bold != null)      ? (Boolean)run.RunProperties.Bold   : textStyle.Bold;
                    textStyle.Italic    = (run.RunProperties.Italic != null)    ? (Boolean)run.RunProperties.Italic : textStyle.Italic;
                    textStyle.Underline = (run.RunProperties.Underline != null) ? true                              : textStyle.Underline;

                    sceneObject.StyleList.Add(textStyle);

                    textFragment.StyleId = sceneObject.StyleList.IndexOf(textStyle);

                    //Calculate the internal text fragments position (inside the scene object)
                    var xy = CalculateInternalTextFragmentPositions(textFragment, 
                                                                    textStyle,
                                                                    simpleSceneObject.ClipWidth, 
                                                                    simpleSceneObject.ClipHeight);                    
                    textFragment.X = xy.Item1;
                    textFragment.Y = xy.Item2;                 

                    sceneObject.FragmentsList.Add(textFragment);
                }

                sceneObjectList.Add(sceneObject);
            }

            return sceneObjectList;
        }

        private Tuple<int,int> CalculateInternalTextFragmentPositions(TextFragment textFragment, TextStyle textStyle, int p1, int p2)
        {
            int x = 0, y = 0;

            return Tuple.Create(x, y);
        }

        
    }
}
