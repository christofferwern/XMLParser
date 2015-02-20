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

            //try
            //{
                // Open the presentation file as read-only.
                using (_presentationDocument = PresentationDocument.Open(_path, false))
                {
                    //Read all colors from theme
                    readColorFromTheme();


                    //Retrive the presentation part
                    var presentation = _presentationDocument.PresentationPart.Presentation;

                    //Get the size of presentation
                    PresentationML.SlideSize slideInfo = presentation.SlideSize;
                    _presentationSizeX = slideInfo.Cx.Value;
                    _presentationSizeY = slideInfo.Cy.Value;

                    foreach (PresentationML.SlideMasterId masterId in presentation.SlideMasterIdList)
                    {
                        SlideMasterPart masterPart = _presentationDocument.PresentationPart.GetPartById(masterId.RelationshipId) as SlideMasterPart;
                        PresentationML.SlideMaster slideMaster = masterPart.SlideMaster;
                        if (slideMaster.CommonSlideData.FirstChild.GetType() == typeof(PresentationML.Background))
                        {
                            foreach(PresentationML.BackgroundProperties bgP in slideMaster.Descendants<PresentationML.BackgroundProperties>())
                            {

                                ShapeObject shapeObject;
                                getBackgroundShape(bgP, out shapeObject);

                                if (shapeObject != null)
                                    _presentationObject.BackgroundSceneObjectList.Add(shapeObject);

                            }
                        }
                    }

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

                    _presentationObject.ConvertToYoobaUnits(_presentationSizeX, _presentationSizeY);
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

            //Check for each backgroundproperty in each slide
            foreach (PresentationML.BackgroundProperties bgP in slide.Descendants<PresentationML.BackgroundProperties>())
            {
                //If no background tag, go to next iteration
                if (bgP == null)
                    continue;
                ShapeObject shapeObject;
                getBackgroundShape(bgP, out shapeObject);

                if (shapeObject != null)
                    backgroundShapelist.Add(shapeObject);

            }

            return backgroundShapelist;
        }

        private void getBackgroundShape(PresentationML.BackgroundProperties bg, out ShapeObject shapeObject)
        {
            SimpleSceneObject simpleSceneObject = new SimpleSceneObject();
            simpleSceneObject.ClipHeight = _presentationSizeY;
            simpleSceneObject.ClipWidth = _presentationSizeX;

            //string slide_name = splitUriToImageName(slidePart.Uri); //Get the slide for this background

            OpenXmlElement bg_type = bg.FirstChild;
            shapeObject = null;
            //Check if first child in background tag is a solidFill, gradFill or blipFill
            switch (bg_type.XmlQualifiedName.Name)
            {

                case "solidFill":
                    SolidBackground solidBg = new SolidBackground();
                    shapeObject = new ShapeObject(simpleSceneObject, ShapeObject.shape_type.Rectangle);

                    if (bg_type.FirstChild.GetType().Equals(typeof(DrawingML.RgbColorModelHex)))
                        foreach (var value in bg_type.Descendants<DrawingML.RgbColorModelHex>())
                            solidBg.BgColor = value.Val;
                    else
                        foreach (var value in bg_type.Descendants<DrawingML.SchemeColor>())
                            solidBg.BgColor = getColorFromTheme(value.Val);

                    IEnumerator<OpenXmlElement> colorInfo = bg_type.FirstChild.GetEnumerator();

                    //Get the alpha value for each gradient stop
                    while (colorInfo.MoveNext())
                    {
                        ColorConverter conv = new ColorConverter();
                        
                        var type = colorInfo.Current;

                        if (type.GetType() == typeof(DrawingML.Alpha))
                        {
                            var Alpha = type as DrawingML.Alpha;
                            solidBg.Alpha = Alpha.Val;
                        }
                        if (type.GetType() == typeof(DrawingML.SaturationModulation))
                        {
                            var SaturationModulation = type as DrawingML.SaturationModulation;
                            solidBg.BgColor = conv.SetSaturation(solidBg.Color, SaturationModulation.Val);
                        }
                        if (type.GetType() == typeof(DrawingML.Shade))
                        {
                            var Shade = type as DrawingML.Shade;
                            solidBg.BgColor = conv.SetTint(solidBg.Color, Shade.Val);
                        }
                        if (type.GetType() == typeof(DrawingML.Tint))
                        {
                            var Tint = type as DrawingML.Tint;
                            solidBg.BgColor = conv.SetShade(solidBg.Color, Tint.Val);
                        }
                        if (type.GetType() == typeof(DrawingML.LuminanceModulation))
                        {
                            var LuminanceModulation = type as DrawingML.LuminanceModulation;
                            solidBg.BgColor = conv.SetBrightness(solidBg.Color, LuminanceModulation.Val);
                        }

                    }

                    shapeObject.FillColor = solidBg.BgColor;
                    shapeObject.FillAlpha = solidBg.Alpha;
                    shapeObject.GradientType = solidBg.BackgroundType;

                    break;
                case "gradFill":

                    shapeObject = new ShapeObject(simpleSceneObject, ShapeObject.shape_type.Rectangle);

                    GradientBackground gradientBg = new GradientBackground();

                    bool validGradientType = false;

                    IEnumerator<OpenXmlElement> gradientType = bg_type.GetEnumerator();

                    while (gradientType.MoveNext() && validGradientType == false)
                    {
                        if (gradientType.Current.GetType() == typeof(DrawingML.PathGradientFill))
                        {
                            gradientBg.GradientType = "radial";
                            validGradientType = true;
                        }
                        else if (gradientType.Current.GetType() == typeof(DrawingML.LinearGradientFill))
                        {
                            gradientBg.GradientType = "linear";
                            var angle = gradientType.Current as DrawingML.LinearGradientFill;
                            gradientBg.Angle = angle.Angle;
                            validGradientType = true;
                        }
                        else
                            validGradientType = false;
                    }

                    if (validGradientType)
                    {
                        foreach (DrawingML.GradientStop gs in bg_type.FirstChild.Descendants<DrawingML.GradientStop>())
                        {

                            //Create a new gradientinfo to store color, alpha and position for each gradient stop.
                            GradientInfo gradInfo = new GradientInfo();

                            //Get the position value for each gradient stop.
                            gradInfo.Position = gs.Position.Value;

                            //Get the color value for each gradient stop.
                            if (gs.RgbColorModelHex != null)
                                gradInfo.GradColor = gs.RgbColorModelHex.Val;
                            else
                                gradInfo.GradColor = getColorFromTheme(gs.SchemeColor.Val);

                            //List of all childrens gradientstop's first child.
                            IEnumerator<OpenXmlElement> alpha = gs.FirstChild.GetEnumerator();
                            Console.WriteLine(gradInfo.Color);
                            //Get the alpha value for each gradient stop
                            while (alpha.MoveNext())
                            {
                                ColorConverter conv = new ColorConverter();

                                var colorType = alpha.Current;

                                if (colorType.GetType() == typeof(DrawingML.Alpha))
                                {
                                    var Alpha = colorType as DrawingML.Alpha;
                                    gradInfo.Alpha = Alpha.Val;
                                }
                                if (colorType.GetType() == typeof(DrawingML.SaturationModulation))
                                {
                                    var SaturationModulation = colorType as DrawingML.SaturationModulation;
                                    gradInfo.GradColor = conv.SetSaturation(gradInfo.Color, SaturationModulation.Val);
                                    Console.WriteLine(gradInfo.GradColor);
                                }
                                if (colorType.GetType() == typeof(DrawingML.Shade))
                                {
                                    var Shade = colorType as DrawingML.Shade;
                                    gradInfo.GradColor = conv.SetTint(gradInfo.Color, Shade.Val);
                                    Console.WriteLine(gradInfo.GradColor);
                                }
                                if (colorType.GetType() == typeof(DrawingML.Tint))
                                {
                                    var Tint = colorType as DrawingML.Tint;
                                    gradInfo.GradColor = conv.SetShade(gradInfo.Color, Tint.Val);
                                    Console.WriteLine(gradInfo.GradColor);
                                }
                                if (colorType.GetType() == typeof(DrawingML.LuminanceModulation))
                                {
                                    var LuminanceModulation = colorType as DrawingML.LuminanceModulation;
                                    gradInfo.GradColor = conv.SetBrightness(gradInfo.Color, LuminanceModulation.Val);
                                }

                            }
                            //gradInfo.convert();
                            //Console.WriteLine(gradInfo.toString());
                            //Add gradient information to backgroundlist
                            gradientBg.GradientList.Add(gradInfo);
                        }

                        shapeObject.GradientType = gradientBg.GradientType;
                        shapeObject.FillType = gradientBg.BgType;

                        shapeObject.GradientAlphas[0] = gradientBg.GradientList.First().Alpha;
                        shapeObject.GradientAlphas[1] = gradientBg.GradientList.Last().Alpha;
                        shapeObject.GradientFills[0] = gradientBg.GradientList.First().GradColor;
                        shapeObject.GradientFills[1] = gradientBg.GradientList.Last().GradColor;

                        shapeObject.GradientAngle = gradientBg.Angle;

                    }
                    else
                    {
                        Console.WriteLine("Not valid type");
                        shapeObject = null;
                    }

                    break;
                case "blipFill":
                    Console.WriteLine("blipFill");
                    /*foreach (var blip in bg_type.Descendants<DrawingML.Blip>())
                    {
                        string embeded_id = blip.Embed.Value;
                        ImagePart image_part = (ImagePart)slide.SlidePart.GetPartById(embeded_id);
                        //Console.WriteLine(image_part.Uri.OriginalString);

                    }
                    */
                    shapeObject = null;
                    break;
            }

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

                SimpleSceneObject simpleSceneObjectText = new SimpleSceneObject();
                simpleSceneObjectText.BoundsX = powerPointText.X;
                simpleSceneObjectText.BoundsY = powerPointText.Y;
                simpleSceneObjectText.ClipWidth = powerPointText.Cx;
                simpleSceneObjectText.ClipHeight = powerPointText.Cy;

                SimpleSceneObject simpleSceneObjectShape = new SimpleSceneObject();
                simpleSceneObjectShape.BoundsX = powerPointText.X;
                simpleSceneObjectShape.BoundsY = powerPointText.Y;
                simpleSceneObjectShape.ClipWidth = powerPointText.Cx;
                simpleSceneObjectShape.ClipHeight = powerPointText.Cy;

                //Get info about SimpleSceneObject

                //Get the rotation of scene object
                foreach (var xfrm in sp.Descendants<DrawingML.Transform2D>())
                {
                    simpleSceneObjectText.Rotation = (xfrm.Rotation != null) ? xfrm.Rotation : simpleSceneObjectText.Rotation;
                    simpleSceneObjectShape.Rotation = (xfrm.Rotation != null) ? xfrm.Rotation : simpleSceneObjectText.Rotation;
                }

                //Get the positions of the scene object
                foreach (var off in sp.Descendants<DrawingML.Offset>())
                {
                    simpleSceneObjectText.BoundsX = (off.X != null) ? (int)off.X : simpleSceneObjectText.BoundsX;
                    simpleSceneObjectText.BoundsY = (off.Y != null) ? (int)off.Y : simpleSceneObjectText.BoundsY;
                    simpleSceneObjectShape.BoundsX = (off.X != null) ? (int)off.X : simpleSceneObjectText.BoundsX;
                    simpleSceneObjectShape.BoundsY = (off.Y != null) ? (int)off.Y : simpleSceneObjectText.BoundsY;
                
                }

                //Get the size of the shape object
                foreach (var ext in sp.Descendants<DrawingML.Extents>())
                {
                    simpleSceneObjectText.ClipWidth = (ext.Cx != null) ? (int)ext.Cx : simpleSceneObjectText.ClipWidth;
                    simpleSceneObjectText.ClipHeight = (ext.Cy != null) ? (int)ext.Cy : simpleSceneObjectText.ClipHeight;
                    simpleSceneObjectShape.ClipWidth = (ext.Cx != null) ? (int)ext.Cx : simpleSceneObjectText.ClipWidth;
                    simpleSceneObjectShape.ClipHeight = (ext.Cy != null) ? (int)ext.Cy : simpleSceneObjectText.ClipHeight;
                }

                //Get the alignment of the shape object
                foreach (DrawingML.Paragraph p in sp.Descendants<DrawingML.Paragraph>())
                {
                    foreach(var child in p.ChildElements)
                    {
                        if (child.LocalName == "pPr")
                        {
                            if (p.ParagraphProperties.Alignment != null)
                            {
                                powerPointText.Alignment = p.ParagraphProperties.Alignment.Value.ToString();
                            }
                        }

                    }
                }

                //If textbody contains listinfo
                List<PowerPointText> listStyleList = new List<PowerPointText>();
                if (sp.TextBody.ListStyle != null)
                {
                    if (sp.TextBody.ListStyle.HasChildren)
                    {
                        PowerPointText p1 = GetLevelProperties<DrawingML.Level1ParagraphProperties>(sp.TextBody.ListStyle.Level1ParagraphProperties); p1.Level = 1;
                        PowerPointText p2 = GetLevelProperties<DrawingML.Level2ParagraphProperties>(sp.TextBody.ListStyle.Level2ParagraphProperties); p2.Level = 2;
                        PowerPointText p3 = GetLevelProperties<DrawingML.Level3ParagraphProperties>(sp.TextBody.ListStyle.Level3ParagraphProperties); p3.Level = 3;
                        PowerPointText p4 = GetLevelProperties<DrawingML.Level4ParagraphProperties>(sp.TextBody.ListStyle.Level4ParagraphProperties); p4.Level = 4;
                        PowerPointText p5 = GetLevelProperties<DrawingML.Level5ParagraphProperties>(sp.TextBody.ListStyle.Level5ParagraphProperties); p5.Level = 5;
                        PowerPointText p6 = GetLevelProperties<DrawingML.Level6ParagraphProperties>(sp.TextBody.ListStyle.Level6ParagraphProperties); p6.Level = 6;
                        PowerPointText p7 = GetLevelProperties<DrawingML.Level7ParagraphProperties>(sp.TextBody.ListStyle.Level7ParagraphProperties); p7.Level = 7;
                        PowerPointText p8 = GetLevelProperties<DrawingML.Level8ParagraphProperties>(sp.TextBody.ListStyle.Level8ParagraphProperties); p8.Level = 8;
                        PowerPointText p9 = GetLevelProperties<DrawingML.Level9ParagraphProperties>(sp.TextBody.ListStyle.Level9ParagraphProperties); p9.Level = 9;
                        listStyleList.Add(p1);
                        listStyleList.Add(p2);
                        listStyleList.Add(p3);
                        listStyleList.Add(p4);
                        listStyleList.Add(p5);
                        listStyleList.Add(p6);
                        listStyleList.Add(p7);
                        listStyleList.Add(p8);
                        listStyleList.Add(p9);
                    }
                }

                //Check if shape has background and line properties.
                if (sp.ShapeProperties != null)
                {
                    if(sp.ShapeProperties.HasChildren)
                    {
                        bool writeShape = false;
                        ShapeObject shapeObject = new ShapeObject(simpleSceneObjectShape, ShapeObject.shape_type.Rectangle);
                        foreach(DrawingML.PresetGeometry pre in sp.ShapeProperties.Descendants<DrawingML.PresetGeometry>())
                        {
                            if(pre.Preset == "rect")
                                shapeObject = new ShapeObject(simpleSceneObjectShape, ShapeObject.shape_type.Rectangle);

                            if (pre.Preset == "ellipse")
                                shapeObject = new ShapeObject(simpleSceneObjectShape, ShapeObject.shape_type.Circle);

                            if (pre.Preset == "triangle")
                            {
                                shapeObject.Points = 3;
                                shapeObject = new ShapeObject(simpleSceneObjectShape, ShapeObject.shape_type.Polygon);
                            }
                                
                        }

                        foreach (var child in sp.ShapeProperties)
                        {
                            if (child.LocalName == "solidFill")
                            {
                                DrawingML.SolidFill solidFill = (DrawingML.SolidFill)child;

                                if (solidFill.RgbColorModelHex != null)
                                {
                                    writeShape = true;
                                    shapeObject.FillColor = solidFill.RgbColorModelHex.Val.ToString();
                                }

                                if (solidFill.SchemeColor != null)
                                {
                                    writeShape = true;
                                    shapeObject.FillColor = getColorFromTheme(solidFill.SchemeColor.Val);
                                }
                            }
                        }

                        if (writeShape)
                        {
                            sceneObjectList.Add(shapeObject);
                        }

                    }
                }


                //Add text info to scene object
                TextObject textObject = new TextObject(simpleSceneObjectText);

                int paragraphIndex = 0;
                foreach (DrawingML.Paragraph p in sp.TextBody.Descendants<DrawingML.Paragraph>())
                {
                    PowerPointText temp = new PowerPointText();

                    if (p.ParagraphProperties != null)
                    {
                        if (p.ParagraphProperties.HasAttributes)
                        {
                            if (p.ParagraphProperties.Level!=null)
                                temp.Level = int.Parse(p.ParagraphProperties.Level.InnerText.ToString());
                        }
                    }

                    if (listStyleList.Count > 0)
                    {
                        temp = listStyleList[temp.Level];
                    }  

                    bool HasRun = false;

                    temp.Type = powerPointText.Type;
                    temp.Idx = powerPointText.Idx;

                    temp = GetPowerPointObject(slideLayoutPowerPointShapes, temp);

                    int runIndex = 0;

                    //Go trough all the runs in the text body
                    //they contains the text and some properties
                    foreach (DrawingML.Run run in p.Descendants<DrawingML.Run>())
                    {
                        HasRun = true;
                        //If there is a shape without any text, continue to next shape
                        //if (run.Text.Text.Equals(""))
                        //    continue;

                        TextFragment textFragment = new TextFragment();

                        if (runIndex == 0)
                        {
                            textFragment.Level = temp.Level;

                            if(paragraphIndex!=0)
                                textFragment.NewParagraph = true;
                        }

                        textFragment.setXMLDocumentRoot(ref _rootXmlDoc);
                        textFragment.Text = run.Text.Text;

                        TextStyle textStyle = new TextStyle();

                        textStyle.setXMLDocumentRoot(ref _rootXmlDoc);

                        textStyle.Bold = temp.Bold;
                        textStyle.Italic = temp.Italic;
                        textStyle.Underline = temp.Underline;
                        textStyle.Font = temp.Font;
                        textStyle.FontColor = temp.FontColor;
                        textStyle.FontSize = temp.FontSize;

                        //Get font
                        foreach (var latinFont in run.Descendants<DrawingML.LatinFont>())
                        {
                            textStyle.Font = getFontFromTheme(latinFont.Typeface);
                        }

                        //If font is empty, set it to default font
                        if (textStyle.Font == "")
                            textStyle.Font = getFontFromTheme("+mj-lt");


                        if (run.RunProperties != null)
                        {
                            //Get run properties (size, bold, italic, underline) and insert into style
                            textStyle.FontSize = (run.RunProperties.FontSize != null) ? (int)run.RunProperties.FontSize : textStyle.FontSize;
                            textStyle.Bold = (run.RunProperties.Bold != null) ? (Boolean)run.RunProperties.Bold : textStyle.Bold;
                            textStyle.Italic = (run.RunProperties.Italic != null) ? (Boolean)run.RunProperties.Italic : textStyle.Italic;
                            textStyle.Underline = (run.RunProperties.Underline != null) ? true : textStyle.Underline;

                            //Get font color
                            if (run.RunProperties.HasChildren)
                            {
                                foreach (var child in run.RunProperties.ChildElements)
                                {
                                    if (child.LocalName == "solidFill")
                                    {
                                        DrawingML.SolidFill soldiFill = (DrawingML.SolidFill)child;

                                        foreach (DrawingML.RgbColorModelHex color in soldiFill.Descendants<DrawingML.RgbColorModelHex>())
                                            textStyle.FontColor = color.Val;

                                        foreach (DrawingML.SchemeColor color in soldiFill.Descendants<DrawingML.SchemeColor>())
                                            textStyle.FontColor = getColorFromTheme(color.Val);

                                    }
                                }
                            }
                        
                        }


                        //If font size still is 0, change it to default size
                        if (textStyle.FontSize == 0)
                        {
                            var textStyles = _presentationDocument.PresentationPart.SlideMasterParts.ElementAt(0).SlideMaster.TextStyles;

                            foreach (DrawingML.DefaultRunProperties defRPR in textStyles.BodyStyle.Level1ParagraphProperties.Descendants<DrawingML.DefaultRunProperties>())
                            {
                                if (defRPR.FontSize != null)
                                {
                                    textStyle.FontSize = defRPR.FontSize;
                                }
                            }
                        }

                        //If font color still is empty, set it to default color
                        if (textStyle.FontColor == "")
                        {
                            textStyle.FontColor = "000000";
                        }

                        //Add textStyle to StyleList of the sceneObject
                        textObject.addToStyleList(textStyle);

                        //Iteratates through StyleList of sceneObject for items that is equal to textstyle and returns the index of that style
                        if (textObject.StyleList.IndexOf(textStyle) == -1)
                        {
                            foreach (TextStyle item in textObject.StyleList)
                                if (textStyle.isEqual(item))
                                    textFragment.StyleId = textObject.StyleList.IndexOf(item);
                        }
                        else
                            textFragment.StyleId = textObject.StyleList.IndexOf(textStyle);

                        //Calculate the internal text fragments position (inside the scene object)
                        var xy = CalculateInternalTextFragmentPositions(textFragment,
                                                                        textStyle,
                                                                        simpleSceneObjectText.ClipWidth,
                                                                        simpleSceneObjectText.ClipHeight);
                        textFragment.X = xy.Item1;
                        textFragment.Y = xy.Item2;

                        textObject.FragmentsList.Add(textFragment);

                        runIndex++;
                    }

                    
                    //If paragraph has no run means it has no text, but it still will correspong to a new line 
                    if (!HasRun)
                    {
                        if (textObject.FragmentsList.Count>0)
                            textObject.FragmentsList.Last().Breaks++;
                    }


                    //Increase the paragragh index
                    paragraphIndex++;

                }

                //Put the text object properties to the first style in style list
                if (textObject.StyleList.Count > 0)
                {
                    textObject.Bold = textObject.StyleList[0].Bold;
                    textObject.Italic = textObject.StyleList[0].Italic;
                    textObject.Underline = textObject.StyleList[0].Underline;
                    textObject.Size = textObject.StyleList[0].FontSize;
                    textObject.Color = textObject.StyleList[0].FontColor;
                    textObject.Align = powerPointText.Alignment;
                }

                //If alignment is empty, set it to left
                if (textObject.Align == "")
                    textObject.Align = "left";

                if (textObject.FragmentsList.Count < 1)
                    continue;

                sceneObjectList.Add(textObject);
            }

            return sceneObjectList;
        }


        //Returns a list of all the template shapes in the slide master
        private List<PowerPointText> GetSlideMasterPowerPointShapes(PresentationDocument presentationDocument)
        {
            List<PowerPointText> slideMasterPowerPointShapes = new List<PowerPointText>();

            PresentationML.SlideMaster slideMaster = presentationDocument.PresentationPart.SlideMasterParts.ElementAt(0).SlideMaster;
            PresentationML.ShapeTree shapeTree = slideMaster.CommonSlideData.ShapeTree;

            foreach (PresentationML.Shape sp in shapeTree.Descendants<PresentationML.Shape>())
            {
                List<PowerPointText> powerPointTextList = new List<PowerPointText>();

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
                        PowerPointText p1 = GetLevelProperties<DrawingML.Level1ParagraphProperties>(titleStyle.Level1ParagraphProperties); p1.Level = 0;
                        PowerPointText p2 = GetLevelProperties<DrawingML.Level2ParagraphProperties>(titleStyle.Level2ParagraphProperties); p2.Level = 1;
                        PowerPointText p3 = GetLevelProperties<DrawingML.Level3ParagraphProperties>(titleStyle.Level3ParagraphProperties); p3.Level = 2;
                        PowerPointText p4 = GetLevelProperties<DrawingML.Level4ParagraphProperties>(titleStyle.Level4ParagraphProperties); p4.Level = 3;
                        PowerPointText p5 = GetLevelProperties<DrawingML.Level5ParagraphProperties>(titleStyle.Level5ParagraphProperties); p5.Level = 4;
                        PowerPointText p6 = GetLevelProperties<DrawingML.Level6ParagraphProperties>(titleStyle.Level6ParagraphProperties); p6.Level = 5;
                        PowerPointText p7 = GetLevelProperties<DrawingML.Level7ParagraphProperties>(titleStyle.Level7ParagraphProperties); p7.Level = 6;
                        PowerPointText p8 = GetLevelProperties<DrawingML.Level8ParagraphProperties>(titleStyle.Level8ParagraphProperties); p8.Level = 7;
                        PowerPointText p9 = GetLevelProperties<DrawingML.Level9ParagraphProperties>(titleStyle.Level9ParagraphProperties); p9.Level = 8;

                        powerPointTextList.Add(p1);
                        powerPointTextList.Add(p2);
                        powerPointTextList.Add(p3);
                        powerPointTextList.Add(p4);
                        powerPointTextList.Add(p5);
                        powerPointTextList.Add(p6);
                        powerPointTextList.Add(p7);
                        powerPointTextList.Add(p8);
                        powerPointTextList.Add(p9);
                    }
                }
                //Get the general body attributes 
                else if (powerPointText.Type.Equals("body"))
                {
                    var bodyStyle = slideMaster.TextStyles.BodyStyle;

                    if (bodyStyle.HasChildren)
                    {
                        PowerPointText p1 = GetLevelProperties<DrawingML.Level1ParagraphProperties>(bodyStyle.Level1ParagraphProperties); p1.Level = 0;
                        PowerPointText p2 = GetLevelProperties<DrawingML.Level2ParagraphProperties>(bodyStyle.Level2ParagraphProperties); p2.Level = 1;
                        PowerPointText p3 = GetLevelProperties<DrawingML.Level3ParagraphProperties>(bodyStyle.Level3ParagraphProperties); p3.Level = 2;
                        PowerPointText p4 = GetLevelProperties<DrawingML.Level4ParagraphProperties>(bodyStyle.Level4ParagraphProperties); p4.Level = 3;
                        PowerPointText p5 = GetLevelProperties<DrawingML.Level5ParagraphProperties>(bodyStyle.Level5ParagraphProperties); p5.Level = 4;
                        PowerPointText p6 = GetLevelProperties<DrawingML.Level6ParagraphProperties>(bodyStyle.Level6ParagraphProperties); p6.Level = 5;
                        PowerPointText p7 = GetLevelProperties<DrawingML.Level7ParagraphProperties>(bodyStyle.Level7ParagraphProperties); p7.Level = 6;
                        PowerPointText p8 = GetLevelProperties<DrawingML.Level8ParagraphProperties>(bodyStyle.Level8ParagraphProperties); p8.Level = 7;
                        PowerPointText p9 = GetLevelProperties<DrawingML.Level9ParagraphProperties>(bodyStyle.Level9ParagraphProperties); p9.Level = 8;

                        powerPointTextList.Add(p1);
                        powerPointTextList.Add(p2);
                        powerPointTextList.Add(p3);
                        powerPointTextList.Add(p4);
                        powerPointTextList.Add(p5);
                        powerPointTextList.Add(p6);
                        powerPointTextList.Add(p7);
                        powerPointTextList.Add(p8);
                        powerPointTextList.Add(p9);
                    }
                }
                //Get the general other style attributes
                else
                {
                    var otherStyle = slideMaster.TextStyles.OtherStyle;

                    if (otherStyle.HasChildren)
                    {
                        PowerPointText p1 = GetLevelProperties<DrawingML.Level1ParagraphProperties>(otherStyle.Level1ParagraphProperties); p1.Level = 0;
                        PowerPointText p2 = GetLevelProperties<DrawingML.Level2ParagraphProperties>(otherStyle.Level2ParagraphProperties); p2.Level = 1;
                        PowerPointText p3 = GetLevelProperties<DrawingML.Level3ParagraphProperties>(otherStyle.Level3ParagraphProperties); p3.Level = 2;
                        PowerPointText p4 = GetLevelProperties<DrawingML.Level4ParagraphProperties>(otherStyle.Level4ParagraphProperties); p4.Level = 3;
                        PowerPointText p5 = GetLevelProperties<DrawingML.Level5ParagraphProperties>(otherStyle.Level5ParagraphProperties); p5.Level = 4;
                        PowerPointText p6 = GetLevelProperties<DrawingML.Level6ParagraphProperties>(otherStyle.Level6ParagraphProperties); p6.Level = 5;
                        PowerPointText p7 = GetLevelProperties<DrawingML.Level7ParagraphProperties>(otherStyle.Level7ParagraphProperties); p7.Level = 6;
                        PowerPointText p8 = GetLevelProperties<DrawingML.Level8ParagraphProperties>(otherStyle.Level8ParagraphProperties); p8.Level = 7;
                        PowerPointText p9 = GetLevelProperties<DrawingML.Level9ParagraphProperties>(otherStyle.Level9ParagraphProperties); p9.Level = 8;

                        powerPointTextList.Add(p1);
                        powerPointTextList.Add(p2);
                        powerPointTextList.Add(p3);
                        powerPointTextList.Add(p4);
                        powerPointTextList.Add(p5);
                        powerPointTextList.Add(p6);
                        powerPointTextList.Add(p7);
                        powerPointTextList.Add(p8);
                        powerPointTextList.Add(p9);
                    }
                }

                foreach(PowerPointText p in powerPointTextList){
                    p.Idx = powerPointText.Idx;
                    p.Type = powerPointText.Type;
                }

                //Get the rotation
                if (sp.ShapeProperties.Transform2D.HasAttributes)
                {
                    if (sp.ShapeProperties.Transform2D.Rotation != null)
                    {
                        foreach (PowerPointText p in powerPointTextList)
                            p.Rotation = sp.ShapeProperties.Transform2D.Rotation.Value;
                    }
                }

                //Get the position
                if (sp.ShapeProperties.Transform2D.Offset != null)
                {
                    foreach (PowerPointText p in powerPointTextList)
                    {
                        p.X = (int)sp.ShapeProperties.Transform2D.Offset.X;
                        p.Y = (int)sp.ShapeProperties.Transform2D.Offset.Y;
                    }
                        
                }

                //Get the size
                if (sp.ShapeProperties.Transform2D.Extents != null)
                {
                    foreach (PowerPointText p in powerPointTextList)
                    {
                        p.Cx = (int)sp.ShapeProperties.Transform2D.Extents.Cx;
                        p.Cy = (int)sp.ShapeProperties.Transform2D.Extents.Cy;
                    }
                }

                //Get anchor point
                if (sp.TextBody.BodyProperties.Anchor != null)
                {
                    foreach (PowerPointText p in powerPointTextList)
                    {
                        p.Anchor = sp.TextBody.BodyProperties.Anchor;
                    }
                }                 

                //Get alignment
                foreach (var pPr in sp.TextBody.Descendants<DrawingML.ParagraphProperties>())
                {
                    if (pPr.Alignment != null)
                    {
                        foreach (PowerPointText p in powerPointTextList)
                        {
                            p.Alignment = pPr.Alignment;
                        }
                    }    
                }                        
               
                //Override the attributes with the sp values
                if (sp.TextBody.ListStyle.HasChildren)
                {
                    List<PowerPointText> tempList = new List<PowerPointText>();
                    PowerPointText p1 = GetLevelProperties<DrawingML.Level1ParagraphProperties>(sp.TextBody.ListStyle.Level1ParagraphProperties); p1.Level = 1;
                    PowerPointText p2 = GetLevelProperties<DrawingML.Level2ParagraphProperties>(sp.TextBody.ListStyle.Level2ParagraphProperties); p2.Level = 2;
                    PowerPointText p3 = GetLevelProperties<DrawingML.Level3ParagraphProperties>(sp.TextBody.ListStyle.Level3ParagraphProperties); p3.Level = 3;
                    PowerPointText p4 = GetLevelProperties<DrawingML.Level4ParagraphProperties>(sp.TextBody.ListStyle.Level4ParagraphProperties); p4.Level = 4;
                    PowerPointText p5 = GetLevelProperties<DrawingML.Level5ParagraphProperties>(sp.TextBody.ListStyle.Level5ParagraphProperties); p5.Level = 5;
                    PowerPointText p6 = GetLevelProperties<DrawingML.Level6ParagraphProperties>(sp.TextBody.ListStyle.Level6ParagraphProperties); p6.Level = 6;
                    PowerPointText p7 = GetLevelProperties<DrawingML.Level7ParagraphProperties>(sp.TextBody.ListStyle.Level7ParagraphProperties); p7.Level = 7;
                    PowerPointText p8 = GetLevelProperties<DrawingML.Level8ParagraphProperties>(sp.TextBody.ListStyle.Level8ParagraphProperties); p8.Level = 8;
                    PowerPointText p9 = GetLevelProperties<DrawingML.Level9ParagraphProperties>(sp.TextBody.ListStyle.Level9ParagraphProperties); p9.Level = 9;
                    tempList.Add(p1);
                    tempList.Add(p2);
                    tempList.Add(p3);
                    tempList.Add(p4);
                    tempList.Add(p5);
                    tempList.Add(p6);
                    tempList.Add(p7);
                    tempList.Add(p8);
                    tempList.Add(p9);

                    for (int i = 0; i < powerPointTextList.Count; i++)
                    {
                        powerPointTextList[i].Alignment = (tempList[i].Alignment != "") ? tempList[i].Alignment : powerPointTextList[i].Alignment;
                        powerPointTextList[i].Anchor = (tempList[i].Anchor != "") ? tempList[i].Anchor : powerPointTextList[i].Anchor;
                        //powerPointTextList[i].Bold = (tempList[i].Bold != "") ? tempList[i].Bold : powerPointTextList[i].Bold;
                        powerPointTextList[i].Font = (tempList[i].Font != "") ? tempList[i].Font : powerPointTextList[i].Font;
                        powerPointTextList[i].Cx = (tempList[i].Cx != 0) ? tempList[i].Cx : powerPointTextList[i].Cx;
                        powerPointTextList[i].Cy = (tempList[i].Cy != 0) ? tempList[i].Cy : powerPointTextList[i].Cy;
                        powerPointTextList[i].FontColor = (tempList[i].FontColor != "") ? tempList[i].FontColor : powerPointTextList[i].FontColor;
                        powerPointTextList[i].FontSize = (tempList[i].FontSize != 0) ? tempList[i].FontSize : powerPointTextList[i].FontSize;
                        powerPointTextList[i].FontColor = (tempList[i].FontColor != "") ? tempList[i].FontColor : powerPointTextList[i].FontColor;
                        //powerPointTextList[i].Italic = (tempList[i].Italic != "") ? tempList[i].Italic : powerPointTextList[i].Italic;
                        //powerPointTextList[i].Underline = (tempList[i].Underline != "") ? tempList[i].Underline : powerPointTextList[i].Underline;
                    }
                
                }
                    

                foreach (PowerPointText ppt in powerPointTextList)
                {
                    slideMasterPowerPointShapes.Add(ppt);
                }
            }

            //If list do not contains any template named title or body,
            //add those with just the general attributes for them.
            bool title = false, body = false;
            foreach (PowerPointText p in slideMasterPowerPointShapes)
            {
                if (p.Type == "title")
                    title = true;
                if (p.Type == "body")
                    body = true;
                if (title && body)
                    continue;
            }

            if (!title)
            {
                var titleStyle = slideMaster.TextStyles.TitleStyle;

                if (titleStyle.HasChildren)
                {
                    PowerPointText p1 = GetLevelProperties<DrawingML.Level1ParagraphProperties>(titleStyle.Level1ParagraphProperties); p1.Level = 0; p1.Type = "title";
                    PowerPointText p2 = GetLevelProperties<DrawingML.Level2ParagraphProperties>(titleStyle.Level2ParagraphProperties); p2.Level = 1; p2.Type = "title";
                    PowerPointText p3 = GetLevelProperties<DrawingML.Level3ParagraphProperties>(titleStyle.Level3ParagraphProperties); p3.Level = 2; p3.Type = "title";
                    PowerPointText p4 = GetLevelProperties<DrawingML.Level4ParagraphProperties>(titleStyle.Level4ParagraphProperties); p4.Level = 3; p4.Type = "title";
                    PowerPointText p5 = GetLevelProperties<DrawingML.Level5ParagraphProperties>(titleStyle.Level5ParagraphProperties); p5.Level = 4; p5.Type = "title";
                    PowerPointText p6 = GetLevelProperties<DrawingML.Level6ParagraphProperties>(titleStyle.Level6ParagraphProperties); p6.Level = 5; p6.Type = "title";
                    PowerPointText p7 = GetLevelProperties<DrawingML.Level7ParagraphProperties>(titleStyle.Level7ParagraphProperties); p7.Level = 6; p7.Type = "title";
                    PowerPointText p8 = GetLevelProperties<DrawingML.Level8ParagraphProperties>(titleStyle.Level8ParagraphProperties); p8.Level = 7; p8.Type = "title";
                    PowerPointText p9 = GetLevelProperties<DrawingML.Level9ParagraphProperties>(titleStyle.Level9ParagraphProperties); p9.Level = 8; p9.Type = "title";

                    slideMasterPowerPointShapes.Add(p1);
                    slideMasterPowerPointShapes.Add(p2);
                    slideMasterPowerPointShapes.Add(p3);
                    slideMasterPowerPointShapes.Add(p4);
                    slideMasterPowerPointShapes.Add(p5);
                    slideMasterPowerPointShapes.Add(p6);
                    slideMasterPowerPointShapes.Add(p7);
                    slideMasterPowerPointShapes.Add(p8);
                    slideMasterPowerPointShapes.Add(p9);
                }
            }

            if (!body)
            {
                var bodyStyle = slideMaster.TextStyles.BodyStyle;

                if (bodyStyle.HasChildren)
                {
                    PowerPointText p1 = GetLevelProperties<DrawingML.Level1ParagraphProperties>(bodyStyle.Level1ParagraphProperties); p1.Level = 0; p1.Type = "body";
                    PowerPointText p2 = GetLevelProperties<DrawingML.Level2ParagraphProperties>(bodyStyle.Level2ParagraphProperties); p2.Level = 1; p2.Type = "body";
                    PowerPointText p3 = GetLevelProperties<DrawingML.Level3ParagraphProperties>(bodyStyle.Level3ParagraphProperties); p3.Level = 2; p3.Type = "body";
                    PowerPointText p4 = GetLevelProperties<DrawingML.Level4ParagraphProperties>(bodyStyle.Level4ParagraphProperties); p4.Level = 3; p4.Type = "body";
                    PowerPointText p5 = GetLevelProperties<DrawingML.Level5ParagraphProperties>(bodyStyle.Level5ParagraphProperties); p5.Level = 4; p5.Type = "body";
                    PowerPointText p6 = GetLevelProperties<DrawingML.Level6ParagraphProperties>(bodyStyle.Level6ParagraphProperties); p6.Level = 5; p6.Type = "body";
                    PowerPointText p7 = GetLevelProperties<DrawingML.Level7ParagraphProperties>(bodyStyle.Level7ParagraphProperties); p7.Level = 6; p7.Type = "body";
                    PowerPointText p8 = GetLevelProperties<DrawingML.Level8ParagraphProperties>(bodyStyle.Level8ParagraphProperties); p8.Level = 7; p8.Type = "body";
                    PowerPointText p9 = GetLevelProperties<DrawingML.Level9ParagraphProperties>(bodyStyle.Level9ParagraphProperties); p9.Level = 8; p9.Type = "body";

                    slideMasterPowerPointShapes.Add(p1);
                    slideMasterPowerPointShapes.Add(p2);
                    slideMasterPowerPointShapes.Add(p3);
                    slideMasterPowerPointShapes.Add(p4);
                    slideMasterPowerPointShapes.Add(p5);
                    slideMasterPowerPointShapes.Add(p6);
                    slideMasterPowerPointShapes.Add(p7);
                    slideMasterPowerPointShapes.Add(p8);
                    slideMasterPowerPointShapes.Add(p9);
                }
            }


            //Return list
            return slideMasterPowerPointShapes;
        }

        //Returns a list of all the template shapes in the slide layout
        private List<PowerPointText> GetSlideLayoutPowerPointShapes(SlidePart slidePart)
        {
            List<PowerPointText> slideLayoutPowerPointShapes = new List<PowerPointText>();

            PresentationML.ShapeTree shapeTree = slidePart.SlideLayoutPart.SlideLayout.CommonSlideData.ShapeTree;
            foreach (PresentationML.Shape sp in shapeTree.Descendants<PresentationML.Shape>())
            {
                PowerPointText tempPowerPointText = new PowerPointText();
                List<PowerPointText> powerPointTextList = new List<PowerPointText>();

                //Get the PowerPoint type and index for the shape so it can get the information from both 
                //its slideLayout and from the slideMaster/theme.
                if (sp.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.HasChildren)
                {
                    var placeholder = sp.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.PlaceholderShape;

                    foreach (var attr in placeholder.GetAttributes())
                    {
                        if (attr.LocalName == "idx")
                            tempPowerPointText.Idx = int.Parse(attr.Value.ToString());

                        if (attr.LocalName == "type")
                            tempPowerPointText.Type = attr.Value;
                    }
                }

                for (int i = 0; i < 9; i++)
                {
                    PowerPointText temp = new PowerPointText();
                    temp.Level = i;
                    temp.Idx = tempPowerPointText.Idx;
                    temp.Type = tempPowerPointText.Type;
                    powerPointTextList.Add(temp);
                }

                //Get the values from master
                foreach (PowerPointText p in powerPointTextList)
                {
                    if (p == GetPowerPointObject(_slideMasterPowerPointShapes, p) &&
                        p.Type.Contains("Title"))
                    {
                        PowerPointText temp = new PowerPointText();
                        temp.Type = "title";
                        temp = new PowerPointText(GetPowerPointObject(_slideMasterPowerPointShapes, temp));
                        temp.Type = p.Type;
                        p.setVisualAttribues(temp);
                    }
                    else
                    {
                        PowerPointText temp = new PowerPointText();
                        temp = GetPowerPointObject(_slideMasterPowerPointShapes, p);
                        p.setVisualAttribues(temp);
                    }
                }

                //Get shape properties
                if (sp.ShapeProperties != null) 
                { 
                    if (sp.ShapeProperties.Transform2D!=null)
                    {
                        //Get the rotation
                        if (sp.ShapeProperties.Transform2D.HasAttributes)
                        {
                            if (sp.ShapeProperties.Transform2D.Rotation != null)
                            {
                                foreach (PowerPointText p in powerPointTextList)
                                    p.Rotation = sp.ShapeProperties.Transform2D.Rotation.Value;
                            }
                        }

                        //Get the position
                        if (sp.ShapeProperties.Transform2D.Offset != null)
                        {
                            foreach (PowerPointText p in powerPointTextList)
                            {
                                p.X = (int)sp.ShapeProperties.Transform2D.Offset.X;
                                p.Y = (int)sp.ShapeProperties.Transform2D.Offset.Y;
                            }

                        }

                        //Get the size
                        if (sp.ShapeProperties.Transform2D.Extents != null)
                        {
                            foreach (PowerPointText p in powerPointTextList)
                            {
                                p.Cx = (int)sp.ShapeProperties.Transform2D.Extents.Cx;
                                p.Cy = (int)sp.ShapeProperties.Transform2D.Extents.Cy;
                            }
                        }

                    }
                }

                //Get anchor point
                if (sp.TextBody.BodyProperties.Anchor != null)
                {
                    foreach (PowerPointText p in powerPointTextList)
                    {
                        p.Anchor = sp.TextBody.BodyProperties.Anchor;
                    }
                }

                //Get alignment
                foreach (var pPr in sp.TextBody.Descendants<DrawingML.ParagraphProperties>())
                {
                    if (pPr.Alignment != null)
                    {
                        foreach (PowerPointText p in powerPointTextList)
                        {
                            p.Alignment = pPr.Alignment;
                        }
                    }
                } 

                //Override the attributes with the sp list values
                if (sp.TextBody.ListStyle.HasChildren)
                {
                    List<PowerPointText> tempList = new List<PowerPointText>();
                    PowerPointText p1 = GetLevelProperties<DrawingML.Level1ParagraphProperties>(sp.TextBody.ListStyle.Level1ParagraphProperties); p1.Level = 1;
                    PowerPointText p2 = GetLevelProperties<DrawingML.Level2ParagraphProperties>(sp.TextBody.ListStyle.Level2ParagraphProperties); p2.Level = 2;
                    PowerPointText p3 = GetLevelProperties<DrawingML.Level3ParagraphProperties>(sp.TextBody.ListStyle.Level3ParagraphProperties); p3.Level = 3;
                    PowerPointText p4 = GetLevelProperties<DrawingML.Level4ParagraphProperties>(sp.TextBody.ListStyle.Level4ParagraphProperties); p4.Level = 4;
                    PowerPointText p5 = GetLevelProperties<DrawingML.Level5ParagraphProperties>(sp.TextBody.ListStyle.Level5ParagraphProperties); p5.Level = 5;
                    PowerPointText p6 = GetLevelProperties<DrawingML.Level6ParagraphProperties>(sp.TextBody.ListStyle.Level6ParagraphProperties); p6.Level = 6;
                    PowerPointText p7 = GetLevelProperties<DrawingML.Level7ParagraphProperties>(sp.TextBody.ListStyle.Level7ParagraphProperties); p7.Level = 7;
                    PowerPointText p8 = GetLevelProperties<DrawingML.Level8ParagraphProperties>(sp.TextBody.ListStyle.Level8ParagraphProperties); p8.Level = 8;
                    PowerPointText p9 = GetLevelProperties<DrawingML.Level9ParagraphProperties>(sp.TextBody.ListStyle.Level9ParagraphProperties); p9.Level = 9;
                    tempList.Add(p1);
                    tempList.Add(p2);
                    tempList.Add(p3);
                    tempList.Add(p4);
                    tempList.Add(p5);
                    tempList.Add(p6);
                    tempList.Add(p7);
                    tempList.Add(p8);
                    tempList.Add(p9);

                    for (int i = 0; i < powerPointTextList.Count; i++)
                    {
                        powerPointTextList[i].Alignment = (tempList[i].Alignment != "") ? tempList[i].Alignment : powerPointTextList[i].Alignment;
                        powerPointTextList[i].Anchor = (tempList[i].Anchor != "") ? tempList[i].Anchor : powerPointTextList[i].Anchor;
                        powerPointTextList[i].Bold = tempList[i].Bold;
                        powerPointTextList[i].Font = (tempList[i].Font != "") ? tempList[i].Font : powerPointTextList[i].Font;
                        powerPointTextList[i].Cx = (tempList[i].Cx != 0) ? tempList[i].Cx : powerPointTextList[i].Cx;
                        powerPointTextList[i].Cy = (tempList[i].Cy != 0) ? tempList[i].Cy : powerPointTextList[i].Cy;
                        powerPointTextList[i].X = (tempList[i].Cx != 0) ? tempList[i].X : powerPointTextList[i].X;
                        powerPointTextList[i].Y = (tempList[i].Cy != 0) ? tempList[i].Y : powerPointTextList[i].Y;
                        powerPointTextList[i].FontColor = (tempList[i].FontColor != "") ? tempList[i].FontColor : powerPointTextList[i].FontColor;
                        powerPointTextList[i].FontSize = (tempList[i].FontSize != 0) ? tempList[i].FontSize : powerPointTextList[i].FontSize;
                        powerPointTextList[i].FontColor = (tempList[i].FontColor != "") ? tempList[i].FontColor : powerPointTextList[i].FontColor;
                        powerPointTextList[i].Italic = tempList[i].Italic;
                        //powerPointTextList[i].Underline = (tempList[i].Underline != "") ? tempList[i].Underline : powerPointTextList[i].Underline;
                    }

                }

                //Override attributes with the layout paragrapgh values
                foreach (DrawingML.Paragraph p in sp.TextBody.Descendants<DrawingML.Paragraph>())
                {
                    foreach (DrawingML.Run run in p.Descendants<DrawingML.Run>())
                    {
                        if (run.RunProperties != null)
                        {

                            if (run.RunProperties.Bold != null)
                                foreach (PowerPointText ppt in powerPointTextList)
                                    ppt.Bold = run.RunProperties.Bold;


                            if (run.RunProperties.HasChildren)
                            {
                                foreach (var child in run.RunProperties.ChildElements)
                                {
                                    if (child.LocalName == "solidFill")
                                    {
                                        DrawingML.SolidFill soldiFill = (DrawingML.SolidFill)child;

                                        foreach (DrawingML.RgbColorModelHex color in soldiFill.Descendants<DrawingML.RgbColorModelHex>())
                                            foreach (PowerPointText ppt in powerPointTextList)
                                                ppt.FontColor = color.Val;
     
                                        foreach (DrawingML.SchemeColor color in soldiFill.Descendants<DrawingML.SchemeColor>())
                                            foreach (PowerPointText ppt in powerPointTextList)
                                                ppt.FontColor = getColorFromTheme(color.Val);

                                    }
                                }
                            }


                                    

                        }
                    }
                }

                //Add to list
                foreach (PowerPointText p in powerPointTextList)
                {
                    if (p.isEmpty())
                        continue;

                    slideLayoutPowerPointShapes.Add(p);
                }
            }

            return slideLayoutPowerPointShapes;
        }

        private PowerPointText GetPowerPointObject(List<PowerPointText> list, PowerPointText powerPointText)
        {
            PowerPointText temp = new PowerPointText();

            foreach (PowerPointText ppt in list)
            {
                if (ppt.Level == powerPointText.Level)
                {
                    if ((ppt.Idx == powerPointText.Idx && ppt.Idx > 0) && (ppt.Type == powerPointText.Type && ppt.Type != ""))
                    {
                        temp = new PowerPointText(ppt);
                        temp.Type = powerPointText.Type;
                        return temp;
                    }
                }
            }

            foreach (PowerPointText ppt in list)
            {
                if (ppt.Level == powerPointText.Level)
                {
                    if (ppt.Idx == powerPointText.Idx && ppt.Idx > 0)
                    {
                        temp = new PowerPointText(ppt);
                        temp.Type = powerPointText.Type;
                        return temp;
                    }

                    if (ppt.Type == powerPointText.Type && ppt.Type != "")
                    {
                        temp = new PowerPointText(ppt);
                        temp.Idx = powerPointText.Idx;
                        return temp;
                    }
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
            var slideMaster = _presentationDocument.PresentationPart.SlideMasterParts.ElementAt(0).SlideMaster;

            var colorScheme = slideMaster.SlideMasterPart.ThemePart.Theme.ThemeElements.ColorScheme;
            var colorMap = slideMaster.ColorMap;

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
                if (color.Equals(_themeColors.GetValue(i, 0).ToString()))
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


        public PowerPointText GetLevelProperties<T>(T input)
        {
            
            PowerPointText powerPointText = new PowerPointText();

            if (input == null)
                return powerPointText;

            bool hasAttr = (bool) input.GetType().GetMethod("get_HasAttributes").Invoke(input,null);

            if (hasAttr)
            {
                List<OpenXmlAttribute> attrs = (List<OpenXmlAttribute>)input.GetType().GetMethod("GetAttributes").Invoke(input, null);

                foreach (var attr in attrs)
                {
                    if (attr.LocalName == "algn")
                    {
                        powerPointText.Alignment = input.GetType().GetMethod("get_Alignment").Invoke(input, null).ToString();
                    }
                }
            }

            OpenXmlElementList childElements = (OpenXmlElementList) input.GetType().GetMethod("get_ChildElements").Invoke(input, null);

            foreach (var childElement in childElements)
            {
                if (childElement.LocalName == "defRPr")
                {
                    DrawingML.DefaultRunProperties defRPR = (DrawingML.DefaultRunProperties)childElement;

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

                    foreach (var defRPrChild in childElement)
                    {
                        if (defRPrChild.LocalName == "solidFill")
                        {
                            DrawingML.SolidFill solidFill = (DrawingML.SolidFill)defRPrChild;

                            //Get font color
                            foreach (DrawingML.SchemeColor color in solidFill.Descendants<DrawingML.SchemeColor>())
                                powerPointText.FontColor = getColorFromTheme(color.Val);
                            
                            foreach (DrawingML.RgbColorModelHex color in solidFill.Descendants<DrawingML.RgbColorModelHex>())
                                powerPointText.FontColor = color.Val;
                            

                        }
                    }
                }
            }
 
            return powerPointText;
        }

        public int PresentationSizeY
        {
            get { return _presentationSizeY; }
            set { _presentationSizeY = value; }
        }

        public int PresentationSizeX
        {
            get { return _presentationSizeX; }
            set { _presentationSizeX = value; }
        }

    }
}
