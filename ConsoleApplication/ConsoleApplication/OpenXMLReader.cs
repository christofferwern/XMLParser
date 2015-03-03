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

        private List<PowerPointText> _placeHolderListMaster;
        private List<PowerPointText> _placeHolderListLayout;
        private bool _masterLevel, _layoutLevel, _slideLevel;

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

            _placeHolderListMaster = new List<PowerPointText>();
            _placeHolderListLayout = new List<PowerPointText>();

            _masterLevel = false;
            _layoutLevel = false;
            _slideLevel = false;
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
                //Read all colors from theme
                readTheme();

                //Retrive the presentation part
                var presentation = _presentationDocument.PresentationPart.Presentation;

                //Get the size of presentation
                PresentationML.SlideSize slideInfo = presentation.SlideSize;
                _presentationSizeX = slideInfo.Cx.Value;
                _presentationSizeY = slideInfo.Cy.Value;


                //Get the slidemaster scene object, background etc
                PresentationML.SlideMaster slideMaster = _presentationDocument.PresentationPart.SlideMasterParts.ElementAt(0).SlideMaster; ;

                _masterLevel = true;
                foreach (var child in slideMaster.CommonSlideData.ChildElements)
                {
                    if (child.LocalName == "bg")
                        _presentationObject.BackgroundSceneObjectList.AddRange(getSceneObjects((PresentationML.Background)child));

                    if (child.LocalName == "spTree")
                        _presentationObject.BackgroundSceneObjectList.AddRange(getSceneObjects((PresentationML.ShapeTree)child));
                }
                _masterLevel = false;

                //If master list do not contains any template named title or body,
                //add those with just the general attributes for them.
                bool title = false, body = false;
                foreach (PowerPointText p in _placeHolderListMaster)
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
                    PowerPointText[] range = getListStyles(slideMaster.TextStyles.TitleStyle);
                    foreach (PowerPointText p in range)
                        p.Type = "title";

                    _placeHolderListMaster.AddRange(range);
                }
                    
                if (!body)
                {
                    PowerPointText[] range = getListStyles(slideMaster.TextStyles.BodyStyle);
                    foreach (PowerPointText p in range)
                        p.Type = "body";

                    _placeHolderListMaster.AddRange(range);
                }

                //Counter of scenes
                int sceneCounter = 1;

                //Go through all Slides in the PowerPoint presentation
                foreach (PresentationML.SlideId slideID in presentation.SlideIdList)
                {
                    SlidePart slidePart = _presentationDocument.PresentationPart.GetPartById(slideID.RelationshipId) as SlidePart;
                    Scene scene = new Scene(sceneCounter);

                    //Go through all elements in the slide layout
                    _placeHolderListLayout.Clear();
                    _layoutLevel = true;
                    foreach (var child in slidePart.SlideLayoutPart.SlideLayout.CommonSlideData.ChildElements)
                    {
                        if (child.LocalName == "bg")
                            scene.SceneObjectList.AddRange(getSceneObjects((PresentationML.Background)child));

                        if (child.LocalName == "spTree")
                            scene.SceneObjectList.AddRange(getSceneObjects((PresentationML.ShapeTree)child));
                    }
                    _layoutLevel = false;

                    //Go through all elements in the slide
                    _slideLevel = true;
                    foreach (var child in slidePart.Slide.CommonSlideData.ChildElements)
                    {

                        if (child.LocalName == "bg")
                            scene.SceneObjectList.InsertRange(0, getSceneObjects((PresentationML.Background)child));

                        if (child.LocalName == "spTree")
                        {
                            scene.SceneObjectList.AddRange(getSceneObjects((PresentationML.ShapeTree)child));
                        }
                    }
                    _slideLevel = false;


                    //if (sceneCounter == 1)
                    //    foreach (PowerPointText p in _placeHolderListMaster)
                    //        if (p.Level < 1)
                    //            Console.WriteLine(p.toString());

                    //if (sceneCounter == 1)
                    //    foreach (PowerPointText p in _placeHolderListLayout)
                    //        if (p.Level < 1)
                    //            Console.WriteLine(p.toString());

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

            foreach (var child in groupShape.ChildElements)
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
            SimpleSceneObject shapeSimpleSceneObject = new SimpleSceneObject(),
                              textSimpleSceneObject = new SimpleSceneObject();

            ShapeObject shapeObject = new ShapeObject(shapeSimpleSceneObject, ShapeObject.shape_type.Rectangle);
            TextObject textObject = new TextObject(textSimpleSceneObject);
            PowerPointText powerPointText = new PowerPointText();
            const int styles = 9;
            PowerPointText[]    powerPointLevelList = new PowerPointText[styles],
                                listStyleList = new PowerPointText[styles];

            for (int i = 0; i < styles; i++)
            {
                PowerPointText p = new PowerPointText();
                powerPointLevelList[i] = p;

                PowerPointText p2 = new PowerPointText();
                listStyleList[i] = p2;
            }

            //**TODO**
            //Get siblings and check for offsets!!!
            /*
             *Hämta syskon grSpPr -> xfrm 
             *Om inte noll!
             */

            bool HasBg = false, HasLine = false, HasValidGeometry = false, HasText = false, HasTransform = false,
                 solidFillColorChange = false, lineColorChange = false, gradientColorChange = false;

            foreach (var child in shape.ChildElements)
            {
                if(child.LocalName == "nvSpPr")
                {
                    foreach (var nvSpPrChild in child)
                    {
                        if (nvSpPrChild.LocalName == "nvPr")
                        {
                            PresentationML.ApplicationNonVisualDrawingProperties nvPr = (PresentationML.ApplicationNonVisualDrawingProperties)nvSpPrChild;

                            if (nvPr.PlaceholderShape != null)
                            {
                                powerPointText.Idx = (nvPr.PlaceholderShape.Index != null) ? (int)nvPr.PlaceholderShape.Index.Value : powerPointText.Idx;
                                powerPointText.Type = (nvPr.PlaceholderShape.Type != null) ? nvPr.PlaceholderShape.Type.InnerText.ToString() : powerPointText.Type;
                            }
                        }
                    }
                }

                if(child.LocalName == "spPr")
                {
                    PresentationML.ShapeProperties spPr = (PresentationML.ShapeProperties)child;

                    if (spPr.Transform2D != null)
                    {
                        //Dessa ska bero på gruppens Transform2D
                        shapeSimpleSceneObject.BoundsX = (spPr.Transform2D.Offset.X != null) ? (int)spPr.Transform2D.Offset.X : shapeSimpleSceneObject.BoundsX;
                        shapeSimpleSceneObject.BoundsY = (spPr.Transform2D.Offset.Y != null) ? (int)spPr.Transform2D.Offset.Y : shapeSimpleSceneObject.BoundsY;
                        shapeSimpleSceneObject.ClipWidth = (spPr.Transform2D.Extents.Cx != null) ? (int)spPr.Transform2D.Extents.Cx : shapeSimpleSceneObject.ClipWidth;
                        shapeSimpleSceneObject.ClipHeight = (spPr.Transform2D.Extents.Cy != null) ? (int)spPr.Transform2D.Extents.Cy : shapeSimpleSceneObject.ClipHeight;

                        textSimpleSceneObject.BoundsX = (spPr.Transform2D.Offset.X != null) ? (int)spPr.Transform2D.Offset.X : shapeSimpleSceneObject.BoundsX;
                        textSimpleSceneObject.BoundsY = (spPr.Transform2D.Offset.Y != null) ? (int)spPr.Transform2D.Offset.Y : shapeSimpleSceneObject.BoundsY;
                        textSimpleSceneObject.ClipWidth = (spPr.Transform2D.Extents.Cx != null) ? (int)spPr.Transform2D.Extents.Cx : shapeSimpleSceneObject.ClipWidth;
                        textSimpleSceneObject.ClipHeight = (spPr.Transform2D.Extents.Cy != null) ? (int)spPr.Transform2D.Extents.Cy : shapeSimpleSceneObject.ClipHeight;

                        powerPointText.X = (spPr.Transform2D.Offset.X != null) ? (int)spPr.Transform2D.Offset.X : powerPointText.X;
                        powerPointText.Y = (spPr.Transform2D.Offset.Y != null) ? (int)spPr.Transform2D.Offset.Y : powerPointText.Y;
                        powerPointText.Cx = (spPr.Transform2D.Extents.Cx != null) ? (int)spPr.Transform2D.Extents.Cx : powerPointText.Cx;
                        powerPointText.Cy = (spPr.Transform2D.Extents.Cy != null) ? (int)spPr.Transform2D.Extents.Cy : powerPointText.Cy;

                        HasTransform = true;
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
                                shapeObject = new ShapeObject(shapeSimpleSceneObject, ShapeObject.shape_type.Rectangle);
                            }


                            if (prstGeom.Preset == "ellipse" || prstGeom.Preset == "flowChartConnector")
                            {
                                shapeObject = new ShapeObject(shapeSimpleSceneObject, ShapeObject.shape_type.Circle);
                                HasValidGeometry = true;
                            }

                            if (prstGeom.Preset == "triangle" || prstGeom.Preset == "diamond" || prstGeom.Preset == "flowChartDecision" || prstGeom.Preset == "flowChartExtract" ||
                                prstGeom.Preset == "pentagon" || prstGeom.Preset == "hexagon" || prstGeom.Preset == "heptagon" || prstGeom.Preset == "flowChartPreparation" ||
                                prstGeom.Preset == "octagon" || prstGeom.Preset == "decagon" || prstGeom.Preset == "dodecagon" || prstGeom.Preset == "flowChartMerge")
                            {
                                HasValidGeometry = true;

                                if (prstGeom.Preset == "flowChartMerge")
                                {
                                    shapeSimpleSceneObject.Rotation = 180 * 60000;
                                    shapeObject.Rotation += 180 * 60000;
                                }

                                shapeObject = new ShapeObject(shapeSimpleSceneObject, ShapeObject.shape_type.Polygon);
                                shapeObject.Points = (prstGeom.Preset == "triangle" || prstGeom.Preset == "flowChartExtract" || prstGeom.Preset == "flowChartMerge") ? 3 :
                                                     (prstGeom.Preset == "diamond" || prstGeom.Preset == "flowChartDecision") ? 4 :
                                                     (prstGeom.Preset == "pentagon") ? 5 :
                                                     (prstGeom.Preset == "hexagon" || prstGeom.Preset == "flowChartPreparation") ? 6 :
                                                     (prstGeom.Preset == "heptagon") ? 7 :
                                                     (prstGeom.Preset == "octagon") ? 8 :
                                                     (prstGeom.Preset == "decagon") ? 10 :
                                                     12;
                            }

                            if (prstGeom.Preset == "roundRect" || prstGeom.Preset == "flowChartAlternateProcess" || prstGeom.Preset == "flowChartTerminator")
                            {
                                HasValidGeometry = true;
                                shapeObject = new ShapeObject(shapeSimpleSceneObject, ShapeObject.shape_type.Rectangle);
                                shapeObject.CornerRadius = 3906;

                                if (prstGeom.AdjustValueList != null)
                                {
                                    if (prstGeom.AdjustValueList.HasChildren)
                                    {
                                        foreach (var adjChild in prstGeom.AdjustValueList)
                                        {
                                            if (adjChild.LocalName == "gd")
                                            {
                                                DrawingML.ShapeGuide gd = (DrawingML.ShapeGuide)adjChild;

                                                shapeObject.CornerRadius = float.Parse(gd.Formula.Value.ToString().Remove(0, 4));
                                            }
                                        }
                                    }
                                }
                            }

                        }

                        if(spPrChild.LocalName == "solidFill")
                        {
                            HasBg = true;
                            solidFillColorChange = true;
                            shapeObject.FillColor = getColor(spPrChild);
                            shapeObject.FillAlpha = getAlpha(spPrChild);;
                        }   

                        if(spPrChild.LocalName == "gradFill")
                        {
                            shapeObject.GradientType = getGradientType((DrawingML.GradientFill)spPrChild);
                            shapeObject.GradientAngle = getGradientAngle((DrawingML.GradientFill)spPrChild);

                            HasBg = true;
                            List<string> colors = getColors(spPrChild);
                            shapeObject.FillColor1 = colors[0];
                            shapeObject.FillColor2 = colors[colors.Count-1];
                            gradientColorChange = true;
                        }  

                        if(spPrChild.LocalName == "noFill")
                        {
                        
                        } 

                        if(spPrChild.LocalName == "blipFill")
                        {
                        
                        }  

                        if(spPrChild.LocalName == "ln")
                        {
                            DrawingML.Outline ln = (DrawingML.Outline)spPrChild; 
                            
                            shapeObject.LineSize = (ln.Width!=null) ? ln.Width.Value : shapeObject.LineSize;
                            shapeObject.LineEnabled = true;
                            foreach (var lnChild in ln)
                            {
                                if (lnChild.LocalName == "solidFill")
                                {
                                    HasLine = true;
                                    lineColorChange = true;
                                    shapeObject.LineColor = getColor(lnChild);
                                }

                                if (lnChild.LocalName == "gradFill")
                                {
                                    HasLine = true;
                                    lineColorChange = true;
                                    shapeObject.LineColor = getColors(lnChild)[0];
                                }
                            }
                            
                        }
                    }
                }
                
                if(child.LocalName == "txBody")
                {
                    PresentationML.TextBody txBody = (PresentationML.TextBody)child;

                    //Get the data from the above layers
                    if ((powerPointText.Idx > 0) || (powerPointText.Type != ""))
                    {    
                        //if master level, get data from the general styles
                        if (_masterLevel)
                        {
                            var textStyles = _presentationDocument.PresentationPart.SlideMasterParts.ElementAt(0).SlideMaster.TextStyles;

                            //Put data into the level list from the corresponding listStyle
                            if (powerPointText.Type == "title")
                                powerPointLevelList = getListStyles(textStyles.TitleStyle);   
                            else if (powerPointText.Type == "body")
                                powerPointLevelList = getListStyles(textStyles.BodyStyle);  
                            else //others
                                powerPointLevelList = getListStyles(textStyles.OtherStyle);

                            for(int i=0;i<9;i++)
                            {
                                powerPointLevelList[i].setVisualAttribues(powerPointText);
                                powerPointLevelList[i].Idx = powerPointText.Idx;
                                powerPointLevelList[i].Type = powerPointText.Type;
                                powerPointLevelList[i].Level = i;
                            }
                        }

                        //if layout level, get data from master level
                        if(_layoutLevel)
                        {
                            for (int i = 0; i < 9; i++)
                            {
                                powerPointLevelList[i].Idx = powerPointText.Idx;
                                powerPointLevelList[i].Type = powerPointText.Type;
                                powerPointLevelList[i].Level = i;

                                //if the type contains the name title, get info from master on "title"
                                if (powerPointLevelList[i].Type.Contains("Title"))
                                {
                                    PowerPointText temp = new PowerPointText();
                                    temp.Type = "title";
                                    temp = getPowerPointObject(_placeHolderListMaster, temp);
                                    if (temp != null)
                                    {
                                        powerPointLevelList[i].Idx = powerPointText.Idx;
                                        powerPointLevelList[i].Type = powerPointText.Type;
                                        powerPointLevelList[i].setVisualAttribues(temp);
                                        powerPointLevelList[i].Level = i;
                                    }
                                }

                                PowerPointText master = getPowerPointObject(_placeHolderListMaster, powerPointLevelList[i]);
                                if (master != null)
                                {
                                    powerPointLevelList[i].setVisualAttribues(master);
                                    powerPointLevelList[i].Level = i;
                                }
                            }
                        }

                        //if slide level, get data from both master and layout level
                        if (_slideLevel)
                        {
                            PowerPointText layout = getPowerPointObject(_placeHolderListLayout, powerPointText);
                            if (layout != null)
                                powerPointText.setVisualAttribues(layout);
                        }

                        //Get the transform properties
                        if (HasTransform)
                        {
                            powerPointText.X = textSimpleSceneObject.BoundsX;
                            powerPointText.Y = textSimpleSceneObject.BoundsY;
                            powerPointText.Cx = textSimpleSceneObject.ClipWidth;
                            powerPointText.Cy = textSimpleSceneObject.ClipHeight;

                            for (int i = 0; i < 9; i++)
                            {
                                powerPointLevelList[i].X = textSimpleSceneObject.BoundsX;
                                powerPointLevelList[i].Y = textSimpleSceneObject.BoundsY;
                                powerPointLevelList[i].Cx = textSimpleSceneObject.ClipWidth;
                                powerPointLevelList[i].Cy = textSimpleSceneObject.ClipHeight;
                            }
                        }
                    }

                    //if textbody contains listinfo, get that information
                    if (txBody.ListStyle != null)
                    {
                        listStyleList = getListStyles(txBody.ListStyle);
                        for (int i = 0; i < 9; i++)
                            powerPointLevelList[i].setVisualAttribues(listStyleList[i]);
                    }

                    textSimpleSceneObject.BoundsX = powerPointText.X;
                    textSimpleSceneObject.BoundsY = powerPointText.Y;
                    textSimpleSceneObject.ClipWidth = powerPointText.Cx;
                    textSimpleSceneObject.ClipHeight = powerPointText.Cy;

                    textObject = new TextObject(textSimpleSceneObject);

                    int paragraphIndex = 0;
                    
                    foreach (DrawingML.Paragraph p in child.Descendants<DrawingML.Paragraph>())
                    {
                        PowerPointText paragraghPPT = new PowerPointText(powerPointText);
                        //Get the paragraph properties
                        if (p.ParagraphProperties != null)
                        {
                            paragraghPPT.Alignment = (p.ParagraphProperties.Alignment != null) ? p.ParagraphProperties.Alignment.Value.ToString() : powerPointText.Alignment;

                            textObject.Align = paragraghPPT.Alignment;

                            if (p.ParagraphProperties.Level != null)
                            {
                                paragraghPPT.Level = p.ParagraphProperties.Level.Value;

                                PowerPointText temp = powerPointLevelList[paragraghPPT.Level];
                                if(temp!=null)
                                    paragraghPPT.setVisualAttribues(temp);
                            }
                            else
                            {
                                paragraghPPT.Level = 0;
                            }
                        }

                        textObject.Align = (paragraghPPT.Alignment != "") ? paragraghPPT.Alignment : textObject.Align;
                        textObject.Color = (paragraghPPT.FontColor != "") ? paragraghPPT.FontColor : textObject.Color;
                        textObject.Size = (paragraghPPT.FontSize > 0) ? paragraghPPT.FontSize : textObject.Size;

                        bool HasRun = false;
                        //Get the run properties
                        int runIndex = 0;
                        foreach (DrawingML.Run r in p.Descendants<DrawingML.Run>())
                        {
                            HasRun = true;

                            if (r.Text.InnerText != "")
                                HasText = true;

                            PowerPointText runPPT = new PowerPointText(paragraghPPT);

                            TextFragment textFragment = new TextFragment();
                            TextStyle textStyle = new TextStyle();

                            if (runIndex == 0)
                            {
                                textFragment.Level = runPPT.Level;

                                if (paragraphIndex != 0)
                                    textFragment.NewParagraph = true;
                            }

                            if (r.RunProperties != null)
                            {
                                DrawingML.RunProperties rPr = r.RunProperties;
                                runPPT.Bold = (rPr.Bold != null) ? rPr.Bold.Value : runPPT.Bold;
                                runPPT.Italic = (rPr.Italic != null) ? rPr.Italic.Value : runPPT.Italic;
                                runPPT.FontSize = (rPr.FontSize != null) ? rPr.FontSize.Value : runPPT.FontSize;
                                runPPT.Underline = (rPr.Underline != null) ? true : runPPT.Underline;

                                //Get font color
                                if (rPr.HasChildren)
                                {
                                    foreach (var rPrChild in rPr.ChildElements)
                                    {
                                        if (rPrChild.LocalName == "solidFill")
                                        {
                                            runPPT.FontColor = getColor(rPrChild);
                                                
                                        }
                                    }
                                }
                            
                            }

                            textStyle.Bold = runPPT.Bold;
                            textStyle.Italic = runPPT.Italic;
                            textStyle.Underline = runPPT.Underline;
                            textStyle.FontSize = runPPT.FontSize;
                            textStyle.FontColor = runPPT.FontColor;
                            textFragment.Text = r.Text.InnerText;

                            //Set to default stuffs
                            var style = _presentationDocument.PresentationPart.SlideMasterParts.ElementAt(0).SlideMaster.TextStyles.OtherStyle;
                            PowerPointText[] defaultPPT = getListStyles(style);

                            if (textStyle.FontSize == 0)
                                textStyle.FontSize = defaultPPT[runPPT.Level].FontSize;

                            if (textStyle.FontColor == "")
                                textStyle.FontColor = defaultPPT[runPPT.Level].FontColor;

                            if (textObject.Align == "")
                                textObject.Align = defaultPPT[runPPT.Level].Alignment;

                            textObject.StyleList.Add(textStyle);
                            textFragment.StyleId = textObject.StyleList.IndexOf(textStyle);
                            textObject.FragmentsList.Add(textFragment);

                            //Increase the run index
                            runIndex++;
                        }

                        //If paragraph has no run means it has no text, but it still will correspong to a new line 
                        if (!HasRun)
                        {
                            if (textObject.FragmentsList.Count > 0)
                                textObject.FragmentsList.Last().Breaks++;
                        }

                        //Increase the paragragh index
                        paragraphIndex++;
                    }


                }
                
                if(child.LocalName == "style")
                {
                    PresentationML.ShapeStyle style = (PresentationML.ShapeStyle)child;

                    foreach (var styleChild in child)
                    {
                        if (styleChild.LocalName == "lnRef")
                        {
                            DrawingML.LineReference lnRef = (DrawingML.LineReference)styleChild;

                            string color = getColor(lnRef);

                            int lineRefIndex = (int)lnRef.Index.Value;

                            DrawingML.LineStyleList lineStyleList = _presentationDocument.PresentationPart.ThemePart.Theme.ThemeElements.FormatScheme.LineStyleList;

                            int lineStyleIndex = 1;
                            foreach (var ln in lineStyleList)
                            {
                                foreach (var lineStyle in ln)
                                {
                                    if (lineRefIndex == lineStyleIndex)
                                    {
                                        if (lineStyle.LocalName == "solidFill")
                                        {
                                            if (!lineColorChange)
                                            {
                                                HasLine = true;
                                                shapeObject.LineColor = getColor(lineStyle, color);
                                                shapeObject.LineEnabled = true;
                                            }
                                        }

                                        if (lineStyle.LocalName == "gradFill")
                                        {
                                            HasLine = true;
                                        }
                                    }
                                }

                                lineStyleIndex++;
                            }
                        }

                        if (styleChild.LocalName == "fillRef")
                        {
                            DrawingML.FillReference fillRef = (DrawingML.FillReference)styleChild;

                            string color = getColor(fillRef);

                            int fillRefIndex = (int)fillRef.Index.Value;

                            DrawingML.FillStyleList fillStyleList = _presentationDocument.PresentationPart.ThemePart.Theme.ThemeElements.FormatScheme.FillStyleList;

                            int fillStyleIndex = 1;
                            foreach (var fillStyle in fillStyleList)
                            {
                                if (fillRefIndex == fillStyleIndex)
                                {
                                    if (fillStyle.LocalName == "solidFill")
                                    {
                                        if (!solidFillColorChange)
                                        {
                                            HasBg = true;
                                            shapeObject.FillColor = getColor(fillStyle, color);
                                        }
                                    }

                                    if (fillStyle.LocalName == "gradFill")
                                    {
                                        if (!gradientColorChange)
                                        {
                                            shapeObject.GradientType = getGradientType((DrawingML.GradientFill)fillStyle);
                                            shapeObject.GradientAngle = getGradientAngle((DrawingML.GradientFill)fillStyle);

                                            List<string> colors = getColors(fillStyle, color);
                                            shapeObject.FillColor1 = colors.First();
                                            shapeObject.FillColor2 = colors.Last();
                                            shapeObject.FillType = "gradient";

                                            HasBg = true;
                                        }
                                    }
                                }

                                fillStyleIndex++;
                            }

                        }

                        if (styleChild.LocalName == "fontRef")
                        {
                            DrawingML.FontReference fontRef = (DrawingML.FontReference)styleChild;

                            string color = getColor(fontRef);
                            textObject.Color = color;
                            powerPointText.FontColor = color;
                        }
                    }

                }

            }

            if (HasValidGeometry && (HasLine || HasBg))
                sceneObjectList.Add(shapeObject);

            if (_slideLevel && HasText)
            {
                sceneObjectList.Add(textObject);
            }
                
            if (_layoutLevel)
            {
                for(int i=0;i<9;i++)
                {
                    if (powerPointLevelList[i].Idx > 0 || powerPointLevelList[i].Type != "")
                    {
                        _placeHolderListLayout.Add(powerPointLevelList[i]);
                    }
                }
            }

            if (_masterLevel)
            {
                for (int i = 0; i < 9; i++)
                {
                    if (powerPointLevelList[i].Idx > 0 || powerPointLevelList[i].Type != "")
                    {
                        _placeHolderListMaster.Add(powerPointLevelList[i]);
                    }
                }
            }
                
        
            return sceneObjectList;
        }

        private PowerPointText getPowerPointObject(List<PowerPointText> list, PowerPointText powerPointText)
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

            return null;
        }

        private List<SceneObject> getSceneObjects(PresentationML.Background background)
        {
            List<SceneObject> sceneObjectList = new List<SceneObject>();

            SimpleSceneObject backgroundObject = new SimpleSceneObject();
            backgroundObject.ClipHeight = _presentationSizeY;
            backgroundObject.ClipWidth = _presentationSizeX;

            ShapeObject shape = new ShapeObject(backgroundObject, ShapeObject.shape_type.Rectangle);

            if (background.FirstChild.GetType() == typeof(PresentationML.BackgroundStyleReference))
            {
                PresentationML.BackgroundStyleReference bgRef = (PresentationML.BackgroundStyleReference)background.FirstChild;
                string color = getColor(bgRef);

                int bgRefIndex = (int)bgRef.Index.Value - 1000;

                DrawingML.BackgroundFillStyleList bgStyleList = _presentationDocument.PresentationPart.ThemePart.Theme.ThemeElements.FormatScheme.BackgroundFillStyleList;

                int counter = 1;

                foreach (var bgStyle in bgStyleList)
                {
                    if (counter == bgRefIndex)
                    {
                        if (bgStyle.LocalName == "solidFill")
                            shape.FillColor = getColor(bgStyle, color);

                        if (bgStyle.LocalName == "gradFill")
                        {
                            shape.GradientType = getGradientType((DrawingML.GradientFill)bgStyle);
                            shape.GradientAngle = getGradientAngle((DrawingML.GradientFill)bgStyle);

                            List<string> colors = getColors(bgStyle, color);
                            shape.FillColor1 = colors.First();
                            shape.FillColor2 = colors.Last();
                            shape.FillType = "gradient";

                        }
                    }
                    counter++;
                }
                

            }

            //FIX ALPHA

            foreach (DrawingML.SolidFill solidFill in background.Descendants<DrawingML.SolidFill>())
            {
                shape.FillType = "solid";
                shape.FillColor = getColor(solidFill);
            }
            foreach (DrawingML.GradientFill gradientFill in background.Descendants<DrawingML.GradientFill>())
            {

                shape.FillType = "gradient";
                shape.GradientType = getGradientType(gradientFill);
                shape.GradientAngle = getGradientAngle(gradientFill);

                List<string> gradientColorList = getColors(gradientFill);

                shape.FillColor1 = gradientColorList.First();
                shape.FillColor2 = gradientColorList.Last();
            }

            sceneObjectList.Add(shape);

            return sceneObjectList;
        }

        private List<string> getColors(OpenXmlElement xmlElement)
        {
            return getColors(xmlElement, "");
        }

        private List<string> getColors(OpenXmlElement xmlElement, string phClr)
        {
            List<string> colors = new List<string>();

            OpenXmlElement elementType = xmlElement.FirstChild;

            foreach (DrawingML.GradientStop gs in elementType.Descendants<DrawingML.GradientStop>())
                colors.Add(getColor(gs,phClr));

            return colors;
        }
       
        private string getColor(OpenXmlElement xmlElement)
        {
            return getColor(xmlElement, "");
        }

        private string getColor(OpenXmlElement xmlElement, string phClr)
        {
            string color = "";

            OpenXmlElement colorType = xmlElement.FirstChild;
            
            if (colorType.GetType().Equals(typeof(DrawingML.RgbColorModelHex))){
                
                foreach (var colorValue in colorType.Parent.Descendants<DrawingML.RgbColorModelHex>())
                {
                    if (colorValue.Val.ToString() == "phClr")
                        color = phClr;
                    else
                        color = colorValue.Val.ToString();
                }
            }
            else
                foreach (var colorValue in colorType.Parent.Descendants<DrawingML.SchemeColor>())
                {
                    if (colorValue.Val.ToString() == "phClr")
                        color = phClr;
                    else
                        color = getColorFromTheme(colorValue.Val.ToString());
                }

            if (colorType.HasChildren)
            {
                string transformedColor = colorTransforms(colorType, color);
                color = transformedColor;
                return color;
            }
            else
                return color;
        }

        //Returns a transformed color
        private string colorTransforms(OpenXmlElement colorInfo, string currentColor)
        {
            PowerPointColor powerPointColor = new PowerPointColor();

            powerPointColor.Color = currentColor;

            foreach (var colorAttr in colorInfo.ChildElements)
            {
                if (colorAttr.LocalName == "tint")
                    powerPointColor.Tint = ((DrawingML.Tint)colorAttr).Val;
                if (colorAttr.LocalName == "shade")
                    powerPointColor.Shade = ((DrawingML.Shade)colorAttr).Val;
                if (colorAttr.LocalName == "alpha")
                    powerPointColor.Alpha = ((DrawingML.Alpha)colorAttr).Val;
                if (colorAttr.LocalName == "alphaOff")
                    powerPointColor.AlphaOff = ((DrawingML.AlphaOffset)colorAttr).Val;
                if (colorAttr.LocalName == "alphaMod")
                    powerPointColor.AlphaMod = ((DrawingML.AlphaModulation)colorAttr).Val;
                if (colorAttr.LocalName == "hueMod")
                    powerPointColor.HueMod = ((DrawingML.HueModulation)colorAttr).Val;
                if (colorAttr.LocalName == "hueOff")
                    powerPointColor.HueOff = ((DrawingML.HueOffset)colorAttr).Val;
                if (colorAttr.LocalName == "sat")
                    powerPointColor.Sat = ((DrawingML.Saturation)colorAttr).Val;
                if (colorAttr.LocalName == "satOff")
                    powerPointColor.SatOff = ((DrawingML.SaturationOffset)colorAttr).Val;
                if (colorAttr.LocalName == "satMod")
                    powerPointColor.SatMod = ((DrawingML.SaturationModulation)colorAttr).Val;
                if (colorAttr.LocalName == "lum")
                    powerPointColor.Lum = ((DrawingML.Luminance)colorAttr).Val;
                if (colorAttr.LocalName == "lumOff")
                    powerPointColor.LumOff = ((DrawingML.LuminanceOffset)colorAttr).Val;
                if (colorAttr.LocalName == "lumMod")
                    powerPointColor.LumMod = ((DrawingML.LuminanceModulation)colorAttr).Val;
                if (colorAttr.LocalName == "red")
                    powerPointColor.Red = ((DrawingML.Red)colorAttr).Val;
                if (colorAttr.LocalName == "redOff")
                    powerPointColor.RedOff = ((DrawingML.RedOffset)colorAttr).Val;
                if (colorAttr.LocalName == "redMod")
                    powerPointColor.RedMod = ((DrawingML.RedModulation)colorAttr).Val;
                if (colorAttr.LocalName == "green")
                    powerPointColor.Green = ((DrawingML.Green)colorAttr).Val;
                if (colorAttr.LocalName == "greenOff")
                    powerPointColor.GreenOff = ((DrawingML.GreenOffset)colorAttr).Val;
                if (colorAttr.LocalName == "greenMod")
                    powerPointColor.GreenMod = ((DrawingML.GreenModulation)colorAttr).Val;
                if (colorAttr.LocalName == "blue")
                    powerPointColor.Blue = ((DrawingML.Blue)colorAttr).Val;
                if (colorAttr.LocalName == "blueOff")
                    powerPointColor.BlueOff = ((DrawingML.BlueOffset)colorAttr).Val;
                if (colorAttr.LocalName == "blueMod")
                    powerPointColor.BlueMod = ((DrawingML.BlueModulation)colorAttr).Val;
                //if (colorAttr.LocalName == "gamma")
                //    powerPointColor.Gamma = ((DrawingML.Gamma)colorAttr).Val;
                //if (colorAttr.LocalName == "invGamma")
                //    powerPointColor.InvGamma = ((DrawingML.InverseGamma)colorAttr).Val;
                //if (colorAttr.LocalName == "comp")
                //    powerPointColor.Comp = ((DrawingML.Complement)colorAttr).Val;
                //if (colorAttr.LocalName == "inv")
                //    powerPointColor.Inv = ((DrawingML.Inverse)colorAttr).Val;
                //if (colorAttr.LocalName == "gray")
                //    powerPointColor.Gray = ((DrawingML.Gray)colorAttr).Val;

            }
            
            return powerPointColor.getAdjustedColor();
        }

        private string getColorFromTheme(string color)
        {
            for (int i = 0; i < _nrOfThemeColors; i++)
            {
                if (color.Equals(_themeColors.GetValue(i, 0).ToString()))
                {
                    return (string)_themeColors.GetValue(i, 1);
                }
            }

            return color;
        }

        //Read all the colors in the theme
        private void readTheme()
        {
            var slideMaster = _presentationDocument.PresentationPart.SlideMasterParts.ElementAt(0).SlideMaster;

            //Handle the colors
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

            //Handle the line attr
            var formatScheme = slideMaster.SlideMasterPart.ThemePart.Theme.ThemeElements.FormatScheme;

            if (formatScheme.LineStyleList.HasChildren)
            {
                foreach (DrawingML.Outline ln in formatScheme.LineStyleList)
                {
                    _themeLines.Add(ln);
                }
            }


        }

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


        public int getGradientAngle(DrawingML.GradientFill gradFill)
        {
            foreach (var lin in gradFill.Descendants<DrawingML.LinearGradientFill>())
                if(lin.Angle!=null)
                    return lin.Angle.Value;
                                        
            return 0;
        }

        public string getGradientType(DrawingML.GradientFill gradFill)
        {
            foreach (var lin in gradFill)
            {
                if (lin.GetType() == typeof(DrawingML.LinearGradientFill)) return "linear";
                if (lin.GetType() == typeof(DrawingML.PathGradientFill)) return "radial";
            }

            //Default
            return "linear";
        }

        public int getAlpha(OpenXmlElement xmlElement)
        {
            foreach (var child in xmlElement.FirstChild)
            {
                if (child.LocalName == "alpha")
                {
                    DrawingML.Alpha alpha = (DrawingML.Alpha)child;

                    Console.WriteLine(alpha.Val.Value);

                    return alpha.Val.Value;
                }
            }

            //Default
            return 1;
        }


        public PowerPointText[] getListStyles(OpenXmlElement listStyleElement)
        {
            int styles = 9;
            PowerPointText[] array = new PowerPointText[styles];

            for (int i = 0; i < styles; i++)
            {
                PowerPointText p = new PowerPointText();
                array[i] = p;
            }

            int level = 0;
            foreach (var child in listStyleElement)
            {
                if (child.LocalName == "defPPr")
                {
                    //TODO
                }
                else
                {
                    PowerPointText p = getLevelProperties(child);
                    p.Level = level;
                    array[level] = p;
                    level++;
                }
            }

            return array;
        }

        public PowerPointText getLevelProperties<T>(T input)
        {

            PowerPointText powerPointText = new PowerPointText();

            if (input == null)
                return powerPointText;

            bool hasAttr = (bool)input.GetType().GetMethod("get_HasAttributes").Invoke(input, null);

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

            OpenXmlElementList childElements = (OpenXmlElementList)input.GetType().GetMethod("get_ChildElements").Invoke(input, null);

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

                    foreach (var defRPrChild in childElement)
                    {
                        if (defRPrChild.LocalName == "solidFill")
                        {
                            powerPointText.FontColor = getColor(defRPrChild);
                        }
                    }
                }
            }

            return powerPointText;
        }

        
    }
}
