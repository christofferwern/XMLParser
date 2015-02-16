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

        private const int _nrOfThemeColors = 16;
        private string[,] _themeColors;
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
                using (_presentationDocument = PresentationDocument.Open(_path, false))
                {
                    //Read all colors from theme
                    readColorFromTheme();

                    //Retrive the presentation part
                    var presentation = _presentationDocument.PresentationPart.Presentation;

                    //Get all styles stored in the slide master, <PowerPoint Type, TextStyle>
                    _slideMasterPowerPointShapes = GetSlideMasterPowerPointShapes(_presentationDocument);

                    //Counter of scene
                    int sceneCounter = 1;

                    //Go through all Slides in the PowerPoint presentation
                    foreach (PresentationML.SlideId slideID in presentation.SlideIdList)
                    {
                        SlidePart slidePart = _presentationDocument.PresentationPart.GetPartById(slideID.RelationshipId) as SlidePart;
                        Scene scene = new Scene(sceneCounter);

                        //Get background
                        List<SceneObject> backgroundShapelist = GetBackgroundShapesFromSlidePart(slidePart);

                        //Get all text shape
                        List<SceneObject> textShapelist = GetTextShapesFromSlidePart(slidePart);

                        //Get all images
                        List<SceneObject> imageShapelist = GetImageShapesFromSlidePart(slidePart);

                        //Get all text shape
                        List<SceneObject> shapelist = GetShapesFromSlidePart(slidePart);

                        //Add all scene object to the scene
                        scene.addSceneObjects(backgroundShapelist);
                        scene.addSceneObjects(textShapelist);
                        scene.addSceneObjects(imageShapelist);
                        scene.addSceneObjects(shapelist);
                        
                        //Add scene to presentation
                        _presentationObject.addScene(scene);

                        sceneCounter++;
                    }
                }
            //}
            //catch
            //{
            //    Console.WriteLine("Error reading: " + _path);
            //}
        }

        private List<SceneObject> GetShapesFromSlidePart(SlidePart slidePart)
        {
            List<SceneObject> shapelist = new List<SceneObject>();

            return shapelist;
        }

        private List<SceneObject> GetBackgroundShapesFromSlidePart(SlidePart slidePart)
        {
            List<SceneObject> backgroundShapelist = new List<SceneObject>();

            //Get the actual slide
            PresentationML.Slide slide = slidePart.Slide;

            Background customBg = new Background();
            //Check for each backgroundproperty in each slide
            foreach (var bg in slide.Descendants<PresentationML.BackgroundProperties>())
            {
                //If no background tag, go to next iteration
                if (bg == null)
                    continue;

                SimpleSceneObject simpleSceneObject = new SimpleSceneObject();

                string slide_name = splitUriToImageName(slidePart.Uri); //Get the slide for this background
                Console.WriteLine(slide_name);

                var bg_type = bg.FirstChild;

                //Check if first child in background tag is a solidFill, gradFill or blipFill
                switch (bg_type.XmlQualifiedName.Name)
                {

                    case "solidFill":
                        Console.WriteLine("solidFill");
                        ShapeObject shapeObject = new ShapeObject(simpleSceneObject);

                        if (bg_type.FirstChild.GetType().Equals(typeof(DrawingML.RgbColorModelHex)))
                        {
                            foreach (var value in bg_type.Descendants<DrawingML.RgbColorModelHex>())
                                customBg.BgColor = int.Parse(value.Val, System.Globalization.NumberStyles.HexNumber);
                            foreach (var alpha in bg_type.Descendants<DrawingML.Alpha>())
                                customBg.Alpha = alpha.Val;
                        }
                        else
                        { //Theme color
                            foreach (var value in bg_type.Descendants<DrawingML.SchemeColor>())
                                customBg.BgColor = int.Parse(getColorFromTheme(value.Val), System.Globalization.NumberStyles.HexNumber);
                            foreach (var value in bg_type.Descendants<DrawingML.Alpha>())
                                customBg.Alpha = value.Val; 
                        }

                        shapeObject.FillColor = customBg.BgColor;
                        shapeObject.FillAlpha = customBg.Alpha;

                        backgroundShapelist.Add(shapeObject);

                        break;
                    case "gradFill":
                        Console.WriteLine("gradFill");
                        List<KeyValuePair<int, string>> gradPosCol = new List<KeyValuePair<int, string>>();
                        foreach (var gs in bg_type.FirstChild.Descendants<DrawingML.GradientStop>())
                        {

                            if (gs.RgbColorModelHex != null) // self-defined color
                                gradPosCol.Add(new KeyValuePair<int, string>(gs.Position.Value, getColorFromTheme(gs.RgbColorModelHex.Val)));
                            else //Theme color
                                gradPosCol.Add(new KeyValuePair<int, string>(gs.Position.Value, getColorFromTheme(gs.SchemeColor.Val)));

                        }
                        foreach (var item in gradPosCol)
                        {
                            Console.WriteLine(item);
                        }

                        break;
                    case "blipFill":
                        Console.WriteLine("blipFill");
                        foreach (var blip in bg_type.Descendants<DrawingML.Blip>())
                        {
                            string embeded_id = blip.Embed.Value;
                            ImagePart image_part = (ImagePart)slide.SlidePart.GetPartById(embeded_id);
                            Console.WriteLine(image_part.Uri.OriginalString);

                        }

                        break;
                }


            }


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

                //Get the PowerPoint type and index for the shape so it can get the information from both 
                //its slideLayout and from the slideMaster/theme.
                if (sp.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.HasChildren)
                {
                    var placeholder = sp.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.PlaceholderShape;

                    foreach (var attr in placeholder.GetAttributes())
                    {
                        if (attr.LocalName == "idx")
                            powerPointText.Idx = int.Parse(attr.Value.ToString());

                        if (attr.LocalName == "type")
                            powerPointText.Type = attr.Value;
                    }
                }

                powerPointText = GetPowerPointObject(slideLayoutPowerPointShapes, powerPointText);

                SimpleSceneObject simpleSceneObject = new SimpleSceneObject();
                simpleSceneObject.BoundsX = powerPointText.X;
                simpleSceneObject.BoundsY = powerPointText.Y;
                simpleSceneObject.ClipWidth = powerPointText.Cx;
                simpleSceneObject.ClipHeight = powerPointText.Cy;

                //Get info about SimpleSceneObject

                //Get the rotation of scene object
                foreach (var xfrm in sp.Descendants<DrawingML.Transform2D>())
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

                    if (powerPointText.FontColor.Length==6)
                        textStyle.FontColor = int.Parse(powerPointText.FontColor, System.Globalization.NumberStyles.HexNumber); 

                    textStyle.FontSize = powerPointText.FontSize;

                    //Get font
                    foreach (var latinFont in run.Descendants<DrawingML.LatinFont>())
                    {
                        textStyle.Font = getFontFromTheme(latinFont.Typeface);
                    }

                    //Get the texy body color, if it has been changed manually in the ppt file.
                    foreach (var color in run.Descendants<DrawingML.RgbColorModelHex>())
                    {
                        //Convert Hexadeciamal color to integer color
                        textStyle.FontColor = int.Parse(color.Val, System.Globalization.NumberStyles.HexNumber);
                    }

                    //Get run properties (size, bold, italic, underline) and insert into style
                    textStyle.FontSize  = (run.RunProperties.FontSize != null)  ? (int)run.RunProperties.FontSize   : textStyle.FontSize;
                    textStyle.Bold      = (run.RunProperties.Bold != null)      ? (Boolean)run.RunProperties.Bold   : textStyle.Bold;
                    textStyle.Italic    = (run.RunProperties.Italic != null)    ? (Boolean)run.RunProperties.Italic : textStyle.Italic;
                    textStyle.Underline = (run.RunProperties.Underline != null) ? true                              : textStyle.Underline;

                    //Add textStyle to StyleList of the sceneObject
                    sceneObject.addToStyleList(textStyle);

                    //Iteratates through StyleList of sceneObject for items that is equal to textstyle and returns the index of that style
                    if (sceneObject.StyleList.IndexOf(textStyle) == -1)
                    {
                        foreach (TextStyle item in sceneObject.StyleList)
                            if (textStyle.isEqual(item))
                                textFragment.StyleId = sceneObject.StyleList.IndexOf(item);
                    }
                    else
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


        //Returns a list of all the template shapes in the slide master
        private List<PowerPointText> GetSlideMasterPowerPointShapes(PresentationDocument presentationDocument)
        {
            //Read generel fonts
            var fontScheme = presentationDocument.PresentationPart.ThemePart.Theme.ThemeElements.FontScheme;
            string majorFont = fontScheme.MajorFont.LatinFont.Typeface,
                   minorFont = fontScheme.MinorFont.LatinFont.Typeface;
            
            List<PowerPointText> slideMasterPowerPointShapes = new List<PowerPointText>();

            PresentationML.SlideMaster slideMaster = presentationDocument.PresentationPart.SlideMasterParts.ElementAt(0).SlideMaster;
            PresentationML.ShapeTree shapeTree = slideMaster.CommonSlideData.ShapeTree;

            foreach (PresentationML.Shape sp in shapeTree.Descendants<PresentationML.Shape>())
            {
                PowerPointText powerPointText = new PowerPointText();

                //Get the PowerPoint type and index for the shape so it can get the information from both 
                //its slideLayout and from the slideMaster/theme.
                if (sp.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.HasChildren)
                {
                    var placeholder = sp.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.PlaceholderShape;

                    foreach (var attr in placeholder.GetAttributes())
                    {
                        if (attr.LocalName == "idx")
                            powerPointText.Idx = int.Parse(attr.Value.ToString());

                        if (attr.LocalName == "type")
                            powerPointText.Type = attr.Value;
                    }
                }

                //Neither type or index exists for the template
                if (powerPointText.Type.Equals("") && powerPointText.Idx<0)
                    continue;

                //Get the general title attributes
                if (powerPointText.Type.Equals("title"))
                {
                    var titleStyle = slideMaster.TextStyles.TitleStyle;

                    if (titleStyle.HasChildren)
                    {
                        if (titleStyle.Level1ParagraphProperties.Alignment != null)
                        {
                            powerPointText.Alignment = titleStyle.Level1ParagraphProperties.Alignment;
                        } 

                        foreach (DrawingML.DefaultRunProperties defRPR in titleStyle.Level1ParagraphProperties.Descendants<DrawingML.DefaultRunProperties>())
                        {
                            //Get run properties (size, bold, italic, underline) and insert into style
                            powerPointText.FontSize = (defRPR.FontSize != null) ? (int)defRPR.FontSize : powerPointText.FontSize;
                            powerPointText.Bold = (defRPR.Bold != null) ? (Boolean)defRPR.Bold : powerPointText.Bold;
                            powerPointText.Italic = (defRPR.Italic != null) ? (Boolean)defRPR.Italic : powerPointText.Italic;
                            //powerPointText.Underline = (defRPR.Underline != null) ? true : powerPointText.Underline;

                            if (defRPR.Underline != null)
                            {
                                if (defRPR.Underline.Value.ToString() == "sng")
                                {
                                    powerPointText.Underline = true;
                                }
                            }

                            //Get font
                            foreach (DrawingML.LatinFont font in defRPR.Descendants<DrawingML.LatinFont>())
                            {
                                powerPointText.Font = getFontFromTheme(font.Typeface);
                            }

                            //Get font color
                            foreach (DrawingML.SchemeColor color in defRPR.Descendants<DrawingML.SchemeColor>())
                            {
                                powerPointText.FontColor = getColorFromTheme(color.Val);
                                powerPointText.FontColor = (powerPointText.FontColor == "") ? color.Val : powerPointText.FontColor;
                            }
                        }
                    }
                }
                //Get the general body attributes 
                else if (powerPointText.Type.Equals("body"))
                {
                    var bodyStyle = slideMaster.TextStyles.BodyStyle;

                    if (bodyStyle.HasChildren)
                    {
                        if (bodyStyle.Level1ParagraphProperties.Alignment != null)
                            powerPointText.Alignment = bodyStyle.Level1ParagraphProperties.Alignment;

                        foreach (DrawingML.DefaultRunProperties defRPR in bodyStyle.Level1ParagraphProperties.Descendants<DrawingML.DefaultRunProperties>())
                        {
                            //Get run properties (size, bold, italic, underline) and insert into style
                            powerPointText.FontSize = (defRPR.FontSize != null) ? (int)defRPR.FontSize : powerPointText.FontSize;
                            powerPointText.Bold = (defRPR.Bold != null) ? (Boolean)defRPR.Bold : powerPointText.Bold;
                            powerPointText.Italic = (defRPR.Italic != null) ? (Boolean)defRPR.Italic : powerPointText.Italic;
                            //powerPointText.Underline = (defRPR.Underline != null) ? true : powerPointText.Underline;

                            if (defRPR.Underline != null)
                            {
                                if (defRPR.Underline.Value.ToString() == "sng")
                                {
                                    powerPointText.Underline = true;
                                }
                            }

                            //Get font
                            foreach (DrawingML.LatinFont font in defRPR.Descendants<DrawingML.LatinFont>())
                            {
                                powerPointText.Font = font.Typeface;

                                if (powerPointText.Font.Equals("+mj-lt"))
                                {
                                    powerPointText.Font = majorFont;
                                }

                                if (powerPointText.Font.Equals("+mn-lt"))
                                {
                                    powerPointText.Font = minorFont;
                                }
                                    
                            }

                            //Get font color
                            foreach (DrawingML.SchemeColor color in defRPR.Descendants<DrawingML.SchemeColor>())
                            {
                                powerPointText.FontColor = getColorFromTheme(color.Val);
                                powerPointText.FontColor = (powerPointText.FontColor == "") ? color.Val : powerPointText.FontColor;
                            }
                        }
                    }
                }
                //Get the general other style attributes
                else
                {
                    var otherStyle = slideMaster.TextStyles.OtherStyle;

                    if (otherStyle.HasChildren)
                    {
                        if (otherStyle.Level1ParagraphProperties.Alignment != null)
                            powerPointText.Alignment = otherStyle.Level1ParagraphProperties.Alignment;



                        foreach (DrawingML.DefaultRunProperties defRPR in otherStyle.Level1ParagraphProperties.Descendants<DrawingML.DefaultRunProperties>())
                        {
                            //Get run properties (size, bold, italic, underline) and insert into style
                            powerPointText.FontSize = (defRPR.FontSize != null) ? (int)defRPR.FontSize : powerPointText.FontSize;
                            powerPointText.Bold = (defRPR.Bold != null) ? (Boolean)defRPR.Bold : powerPointText.Bold;
                            powerPointText.Italic = (defRPR.Italic != null) ? (Boolean)defRPR.Italic : powerPointText.Italic;
                            //powerPointText.Underline = (defRPR.Underline != null) ? true : powerPointText.Underline;

                            if (defRPR.Underline != null)
                            {
                                if (defRPR.Underline.Value.ToString() == "sng")
                                {
                                    powerPointText.Underline = true;
                                }
                            }

                            
                            //Get font
                            foreach (DrawingML.LatinFont font in defRPR.Descendants<DrawingML.LatinFont>())
                            {
                                powerPointText.Font = font.Typeface;

                                if (powerPointText.Font.Equals("+mj-lt"))
                                {
                                    powerPointText.Font = majorFont;
                                }

                                if (powerPointText.Font.Equals("+mn-lt"))
                                {
                                    powerPointText.Font = minorFont;
                                }

                            }

                            //Get font color
                            foreach (DrawingML.SchemeColor color in defRPR.Descendants<DrawingML.SchemeColor>())
                            {
                                powerPointText.FontColor = getColorFromTheme(color.Val);
                                powerPointText.FontColor = (powerPointText.FontColor == "") ? color.Val : powerPointText.FontColor;
                            }
                        }
                    }
                }

                

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
                    if (sp.TextBody.ListStyle.Level1ParagraphProperties.Alignment!=null)
                        powerPointText.Alignment = sp.TextBody.ListStyle.Level1ParagraphProperties.Alignment;

                    foreach (DrawingML.DefaultRunProperties defRPR in sp.TextBody.ListStyle.Level1ParagraphProperties.Descendants<DrawingML.DefaultRunProperties>())
                    {
                        //Get run properties (size, bold, italic, underline) and insert into style
                        powerPointText.FontSize = (defRPR.FontSize != null) ? (int)defRPR.FontSize : powerPointText.FontSize;
                        powerPointText.Bold = (defRPR.Bold != null) ? (Boolean)defRPR.Bold : powerPointText.Bold;
                        powerPointText.Italic = (defRPR.Italic != null) ? (Boolean)defRPR.Italic : powerPointText.Italic;
                        powerPointText.Underline = (defRPR.Underline != null) ? true : powerPointText.Underline;

                        //Get font
                        foreach (DrawingML.LatinFont font in defRPR.Descendants<DrawingML.LatinFont>())
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

        //Returns a list of all the template shapes in the slide layout
        private List<PowerPointText> GetSlideLayoutPowerPointShapes(SlidePart slidePart)
        {
            List<PowerPointText> slideLayoutPowerPointShapes = new List<PowerPointText>();

            PresentationML.ShapeTree shapeTree = slidePart.SlideLayoutPart.SlideLayout.CommonSlideData.ShapeTree;
            foreach (PresentationML.Shape sp in shapeTree.Descendants<PresentationML.Shape>())
            {
                PowerPointText powerPointText = new PowerPointText();

                //Get the PowerPoint type and index for the shape so it can get the information from both 
                //its slideLayout and from the slideMaster/theme.
                if (sp.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.HasChildren)
                {
                    var placeholder = sp.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.PlaceholderShape;

                    foreach (var attr in placeholder.GetAttributes())
                    {
                        if (attr.LocalName == "idx")
                            powerPointText.Idx = int.Parse(attr.Value.ToString());

                        if (attr.LocalName == "type")
                            powerPointText.Type = attr.Value;
                    }
                }

                if (powerPointText == GetPowerPointObject(_slideMasterPowerPointShapes, powerPointText) &&
                    powerPointText.Type.Contains("Title"))
                { 
                        PowerPointText temp = new PowerPointText();
                        temp.Type = "title";
                        temp = new PowerPointText(GetPowerPointObject(_slideMasterPowerPointShapes, temp));
                        temp.Type = powerPointText.Type;
                        powerPointText = temp;
                }
                else
                {
                    powerPointText = GetPowerPointObject(_slideMasterPowerPointShapes, powerPointText);
                }

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
                            powerPointText.Font = getFontFromTheme(font.Typeface);
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

        private PowerPointText GetPowerPointObject(List<PowerPointText> list, PowerPointText powerPointText)
        {
            PowerPointText temp = new PowerPointText();

            foreach (PowerPointText ppt in list)
            {
                if (ppt.Type == powerPointText.Type && ppt.Type!="")
                {
                    temp = ppt;
                    temp.Idx = powerPointText.Idx;
                    return temp;
                }

                if (ppt.Idx == powerPointText.Idx && ppt.Idx>0)
                {
                    temp = ppt;
                    temp.Type = powerPointText.Type;
                    return temp;
                }
            }

            return powerPointText;
        }

        //Get the power point text object in the list that corresponds to the type
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

        //get font from theme, ( inte snyggt alls =/ )
        private string getFontFromTheme(string font)
        {
            var fontScheme = _presentationDocument.PresentationPart.ThemePart.Theme.ThemeElements.FontScheme;

            if (font == "+mj-lt")
                return fontScheme.MajorFont.LatinFont.Typeface;
            else if (font == "+mn-lt")
                return fontScheme.MinorFont.LatinFont.Typeface;
            else
                return font;
        }

        //Read all the colors in the theme
        private void readColorFromTheme()
        {
            var colorScheme = _presentationDocument.PresentationPart.ThemePart.Theme.ThemeElements.ColorScheme;

            var colorMap = _presentationDocument.PresentationPart.SlideMasterParts.ElementAt(0).SlideMaster.ColorMap;

            int index = 0;

            string[] colorRefNames = new String[_nrOfThemeColors]{ "dk1", "lt1", "dk2", "lt2", 
                                       "accent1", "accent2", "accent3","accent4","accent5","accent6",
                                        "hlink", "folHlink", "tx1", "tx2", "bg1", "bg2"};

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

            if (colorExistInTheme(colorMap.Text1.InnerText))
            {
                _themeColors[index, 0] = colorRefNames[index];
                _themeColors[index, 1] = getColorFromTheme(colorMap.Text1.InnerText);
            }
            index++;
            if (colorExistInTheme(colorMap.Text2.InnerText))
            {
                _themeColors[index, 0] = colorRefNames[index];
                _themeColors[index, 1] = getColorFromTheme(colorMap.Text2.InnerText);
            }
            index++;
            if (colorExistInTheme(colorMap.Background1.InnerText))
            {
                _themeColors[index, 0] = colorRefNames[index];
                _themeColors[index, 1] = getColorFromTheme(colorMap.Background1.InnerText);
            }
            index++;
            if (colorExistInTheme(colorMap.Background2.InnerText))
            {
                _themeColors[index, 0] = colorRefNames[index];
                _themeColors[index, 1] = getColorFromTheme(colorMap.Background2.InnerText);
            }
        }

        //Check if color exists in theme
        public bool colorExistInTheme(string color)
        {
            for (int i = 0; i < _nrOfThemeColors; i++)
            {
                if (color == _themeColors.GetValue(i, 0).ToString())
                {
                    return true;
                }
            }
            return false;
        }

        //Get the color from the theme
        private string getColorFromTheme(string color)
        {
            for(int i=0;i<_nrOfThemeColors;i++)
            {
                if (color == _themeColors.GetValue(i, 0).ToString())
                {
                    return (string) _themeColors.GetValue(i, 1);
                }
            }

            return color;
        }

        //Get the name of file
        private string splitUriToImageName(Uri uri)
        {

            string[] image = uri.OriginalString.Split('/');
            string[] image_name = image[3].Split('.');

            return image_name[0];
        }
    }
}
