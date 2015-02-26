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
                    //Read all colors from theme
                    readTheme();

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
                                scene.SceneObjectList.AddRange(getSceneObjects((PresentationML.Background)child));

                            if (child.LocalName == "spTree")
                                scene.SceneObjectList.AddRange(getSceneObjects((PresentationML.ShapeTree)child));
                        }

                        //Go through all elements in the slide
                        foreach(var child in slidePart.Slide.CommonSlideData.ChildElements)
                        {
                            if(child.LocalName == "bg")
                                scene.SceneObjectList.InsertRange(0,getSceneObjects((PresentationML.Background)child));
                            
                            if(child.LocalName == "spTree")
                                scene.SceneObjectList.AddRange(getSceneObjects((PresentationML.ShapeTree)child));
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

            bool HasBg = false, HasLine = false, HasValidGeometry = false,
                 solidFillColorChange = false, lineColorChange = false, gradientColorChange = false;

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

                            if (prstGeom.Preset == "roundRect" || prstGeom.Preset == "flowChartAlternateProcess" || prstGeom.Preset == "flowChartTerminator")
                            {
                                HasValidGeometry = true;
                                shapeObject = new ShapeObject(simpleSceneObject, ShapeObject.shape_type.Rectangle);
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
                            Console.WriteLine(shapeObject.FillColor);
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
                        }
                    }

                }



            }
            if (HasValidGeometry && (HasLine || HasBg))
                sceneObjectList.Add(shapeObject);

            return sceneObjectList;
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
                    if (colorValue.Val.ToString() == "phClr")
                        color = phClr;
                    else
                        color = getColorFromTheme(colorValue.Val.ToString());


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
                    powerPointColor.Lum = ((DrawingML.LuminanceModulation)colorAttr).Val;
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
    }
}
