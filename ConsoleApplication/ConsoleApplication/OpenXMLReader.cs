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

        private string[,] _themeColors;
        private List<PowerPointText> _slideMasterPowerPointShapes;

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

            _themeColors = new string[12, 2];
            _slideMasterPowerPointShapes = new List<PowerPointText>();

            _presentationObject.getXmlDocument(out _rootXmlDoc);

        }

        public string Path
        {
            get { return _path; }
            set { _path = value; }
        }

        public void read()
        {

            //try
            //{
                // Open the presentation file as read-only.
                using (PresentationDocument presentationDocument = PresentationDocument.Open(_path, false))
                {
                    //Read all colors from theme
                    readColorFromTheme(presentationDocument);

                    //Retrive the presentation part
                    var presentation = presentationDocument.PresentationPart.Presentation;

                    //Get all styles stored in the slide master, <PowerPoint Type, TextStyle>
                    _slideMasterPowerPointShapes = GetSlideMasterPowerPointShapes(presentationDocument);

                    //Go through all Slides in the PowerPoint presentation
                    foreach (PresentationML.SlideId slideID in presentation.SlideIdList)
                    {
                        SlidePart slidePart = presentationDocument.PresentationPart.GetPartById(slideID.RelationshipId) as SlidePart;
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
            //}
            //catch
            //{
            //    Console.WriteLine("Error with reading: " + _path);
            //}
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
            //Get all the slide layout power point shapes
            List<PowerPointText> slideLayoutPowerPointShapes = GetSlideLayoutPowerPointShapes(slidePart);

            List<SceneObject> sceneObjectList = new List<SceneObject>();

            //Get the shape tree, <p:spTree>, which contains all shapes in the slide
            var spTree = slidePart.Slide.CommonSlideData.ShapeTree;

            //Traverse through all the shapes in the tree
            foreach (PresentationML.Shape sp in spTree.Descendants<PresentationML.Shape>())
            {
                PowerPointText powerPointText = new PowerPointText();

                //Get the PowerPoint type for the shape so it can get the information from both 
                //its slideLayout and from the slideMaster/theme.
                var nvPr = sp.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties;
                foreach (var ph in nvPr.Descendants<PresentationML.PlaceholderShape>())
                {
                    powerPointText.Type = ph.Type;
                }

                powerPointText = GetPowerPointObjectOnType(slideLayoutPowerPointShapes, powerPointText.Type);

                Console.WriteLine(powerPointText.toString());

                SimpleSceneObject simpleSceneObject = new SimpleSceneObject();
                simpleSceneObject.BoundsX = powerPointText.X;
                simpleSceneObject.BoundsY = powerPointText.Y;
                simpleSceneObject.ClipWidth = powerPointText.Cx;
                simpleSceneObject.ClipHeight = powerPointText.Cy;

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

                    textStyle.Bold = powerPointText.Bold;
                    textStyle.Italic = powerPointText.Italic;
                    textStyle.Underline = powerPointText.Underline;
                    textStyle.Font = powerPointText.Font;
                    textStyle.FontColor = powerPointText.FontColor;
                    textStyle.FontSize = powerPointText.FontSize;
                    

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

        private List<PowerPointText> GetSlideMasterPowerPointShapes(PresentationDocument presentationDocument)
        {
            List<PowerPointText> slideMasterPowerPointShapes = new List<PowerPointText>();

            PresentationML.SlideMaster slideMaster = presentationDocument.PresentationPart.SlideMasterParts.ElementAt(0).SlideMaster;
            PresentationML.ShapeTree shapeTree = slideMaster.CommonSlideData.ShapeTree;

            foreach (PresentationML.Shape sp in shapeTree.Descendants<PresentationML.Shape>())
            {
                PowerPointText powerPointText = new PowerPointText();

                //Get the PowerPoint type for the shape so it can get the information from both 
                //its slideLayout and from the slideMaster/theme.
                var nvPr = sp.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties;
                foreach (var ph in nvPr.Descendants<PresentationML.PlaceholderShape>())
                {
                    powerPointText.Type = ph.Type;
                }

                if (powerPointText.Type.Equals(""))
                    continue;

                //Get the position
                if (sp.ShapeProperties.Transform2D.Offset != null)
                {
                    powerPointText.X = (int)sp.ShapeProperties.Transform2D.Offset.X;
                    powerPointText.Y = (int)sp.ShapeProperties.Transform2D.Offset.Y;
                }

                //Get the size
                if (sp.ShapeProperties.Transform2D.Extents != null)
                {
                    powerPointText.Cx = (int)sp.ShapeProperties.Transform2D.Extents.Cx;
                    powerPointText.Cy = (int)sp.ShapeProperties.Transform2D.Extents.Cy;
                }

                //Get anchor point
                if (sp.TextBody.BodyProperties.Anchor!=null)
                    powerPointText.Anchor = sp.TextBody.BodyProperties.Anchor;

                //Get alignment
                foreach (var pPr in sp.TextBody.Descendants<DrawingML.ParagraphProperties>())
                {
                    if (pPr.Alignment != null)
                        powerPointText.Alignment = pPr.Alignment;
                }

                if (sp.TextBody.ListStyle.HasChildren)
                {
                    foreach (DrawingML.DefaultRunProperties defRPR in sp.TextBody.ListStyle.Level1ParagraphProperties.Descendants<DrawingML.DefaultRunProperties>())
                    {
                        //Get run properties (size, bold, italic, underline) and insert into style
                        powerPointText.FontSize = (defRPR.FontSize != null) ? (int)defRPR.FontSize : powerPointText.FontSize;
                        powerPointText.Bold = (defRPR.Bold != null) ? (Boolean)defRPR.Bold : powerPointText.Bold;
                        powerPointText.Italic = (defRPR.Italic != null) ? (Boolean)defRPR.Italic : powerPointText.Italic;
                        powerPointText.Underline = (defRPR.Underline != null) ? true : powerPointText.Underline;

                        //Get font
                        foreach (DrawingML.SymbolFont font in defRPR.Descendants<DrawingML.SymbolFont>())
                        {
                            powerPointText.Font = font.Typeface;
                        }

                        //Get font color
                        foreach (DrawingML.SchemeColor color in defRPR.Descendants<DrawingML.SchemeColor>())
                        {
                            powerPointText.FontColor = getColorFromTheme(color.Val);
                        }
                    }
                }

                slideMasterPowerPointShapes.Add(powerPointText);
            }

            return slideMasterPowerPointShapes;
        }

        private List<PowerPointText> GetSlideLayoutPowerPointShapes(SlidePart slidePart)
        {
            List<PowerPointText> slideLayoutPowerPointShapes = new List<PowerPointText>();

            PresentationML.ShapeTree shapeTree = slidePart.SlideLayoutPart.SlideLayout.CommonSlideData.ShapeTree;

            foreach (PresentationML.Shape sp in shapeTree.Descendants<PresentationML.Shape>())
            {
                PowerPointText powerPointText = new PowerPointText();

                //Get the PowerPoint type for the shape so it can get the information from both 
                //its slideLayout and from the slideMaster/theme.
                var nvPr = sp.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties;
                foreach (var ph in nvPr.Descendants<PresentationML.PlaceholderShape>())
                {
                    powerPointText.Type = ph.Type;
                }

                if (powerPointText.Type==null)
                    continue;
                if (powerPointText.Type.Equals(""))
                    continue;

                powerPointText = GetPowerPointObjectOnType(_slideMasterPowerPointShapes, powerPointText.Type);

                //Get the position and size
                if (sp.ShapeProperties.Transform2D != null)
                {
                    powerPointText.X = (int)sp.ShapeProperties.Transform2D.Offset.X;
                    powerPointText.Y = (int)sp.ShapeProperties.Transform2D.Offset.Y;
                    powerPointText.Cx = (int)sp.ShapeProperties.Transform2D.Extents.Cx;
                    powerPointText.Cy = (int)sp.ShapeProperties.Transform2D.Extents.Cy;
                }

                //Get anchor point
                if (sp.TextBody.BodyProperties.Anchor != null)
                    powerPointText.Anchor = sp.TextBody.BodyProperties.Anchor;

                //Get alignment
                foreach (var pPr in sp.TextBody.Descendants<DrawingML.ParagraphProperties>())
                {
                    if (pPr.Alignment != null)
                        powerPointText.Alignment = pPr.Alignment;
                }

                if (sp.TextBody.ListStyle.HasChildren)
                {
                    foreach (DrawingML.DefaultRunProperties defRPR in sp.TextBody.ListStyle.Level1ParagraphProperties.Descendants<DrawingML.DefaultRunProperties>())
                    {
                        //Get run properties (size, bold, italic, underline) and insert into style
                        powerPointText.FontSize = (defRPR.FontSize != null) ? (int)defRPR.FontSize : powerPointText.FontSize;
                        powerPointText.Bold = (defRPR.Bold != null) ? (Boolean)defRPR.Bold : powerPointText.Bold;
                        powerPointText.Italic = (defRPR.Italic != null) ? (Boolean)defRPR.Italic : powerPointText.Italic;
                        powerPointText.Underline = (defRPR.Underline != null) ? true : powerPointText.Underline;

                        //Get font
                        foreach (DrawingML.SymbolFont font in defRPR.Descendants<DrawingML.SymbolFont>())
                        {
                            powerPointText.Font = font.Typeface;
                        }

                        //Get font color
                        foreach (DrawingML.SchemeColor color in defRPR.Descendants<DrawingML.SchemeColor>())
                        {
                            powerPointText.FontColor = getColorFromTheme(color.Val);
                        }
                    }
                }

                slideLayoutPowerPointShapes.Add(powerPointText);
            }

            return slideLayoutPowerPointShapes;
        }

        private PowerPointText GetPowerPointObjectOnType(List<PowerPointText> list, string type)
        {
            PowerPointText powerPointText = new PowerPointText();

            foreach(PowerPointText ppt in list)
            {
                if (ppt.Type == type)
                {
                    powerPointText = ppt;
                    return powerPointText;
                }
            }

            powerPointText.Type = type;
            return powerPointText;
        }

        private Tuple<int,int> CalculateInternalTextFragmentPositions(TextFragment textFragment, TextStyle textStyle, int p1, int p2)
        {
            int x = 0, y = 0;

            return Tuple.Create(x, y);
        }

        private void readColorFromTheme(PresentationDocument presentationDocument)
        {
            var colorScheme = presentationDocument.PresentationPart.ThemePart.Theme.ThemeElements.ColorScheme;

            int index = 0;

            string[] colorRefNames = new String[12]{ "dk1", "lt1", "dk2", "lt2", 
                                       "accent1", "accent2", "accent3","accent4","accent5","accent6",
                                        "hlink", "folHlink"};

            foreach (var childElement in colorScheme.ChildElements)
            {
                string color = "";

                foreach (DrawingML.RgbColorModelHex rgbColor in childElement.Descendants<DrawingML.RgbColorModelHex>())
                {
                    color = rgbColor.Val;
                }

                foreach (DrawingML.SystemColor systemColor in childElement.Descendants<DrawingML.SystemColor>())
                {
                    color = systemColor.LastColor;
                }

                _themeColors[index, 0] = colorRefNames[index];
                _themeColors[index, 1] = color;

                index++;
            }
        }

        private int getColorFromTheme(string color)
        {
            for(int i=0;i<12;i++)
            {
                if (color == _themeColors.GetValue(i, 0).ToString())
                {
                    string hexColor = (string) _themeColors.GetValue(i, 1);
                    return int.Parse(hexColor, System.Globalization.NumberStyles.HexNumber);
                }
            }

            return 0;
        }
    }
}
