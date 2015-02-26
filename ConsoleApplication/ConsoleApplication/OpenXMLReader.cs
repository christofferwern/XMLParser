﻿using System;
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

            SimpleSceneObject backgroundObject = new SimpleSceneObject();

            ShapeObject shape = new ShapeObject(backgroundObject, ShapeObject.shape_type.Rectangle);

            //FIX ALPHA

            foreach (DrawingML.SolidFill solidFill in background.Descendants<DrawingML.SolidFill>())
            {
                shape.FillColor = getColor(solidFill);
                Console.WriteLine(shape.FillColor);

            }
            foreach (var gradientFill in background.Descendants<DrawingML.GradientFill>())
            {
                List<string> gradientColorList = getColors(gradientFill);
            }
                



            return sceneObjectList;
        }

        private List<string> getColors(OpenXmlElement xmlElement)
        {
            return getColors(xmlElement, "");
        }
        private List<string> getColors(OpenXmlElement xmlElement, string phClr)
        {
            List<string> colors = new List<string>();

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

    }
}