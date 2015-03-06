using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using System.Runtime.Serialization.Json;
using System.IO;
using SavingImage = System.Drawing.Image;

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
        private List<ImageResponse> _listOfimageResponse;
        private bool _masterLevel, _layoutLevel, _slideLevel;

        private XmlDocument _rootXmlDoc;

        private PresentationML.Slide currentSlide;
        private PresentationML.SlideMaster currentMaster;

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

            _listOfimageResponse = new List<ImageResponse>();

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
                    
                    //Get the image response from Yooba database
                    getImagesFromDatabase();
                    
                    //Retrive the presentation part
                    var presentation = _presentationDocument.PresentationPart.Presentation;

                    //foreach(var part in _presentationDocument.CoreFilePropertiesPart.OpenXmlPackage.Package.PackageProperties.ContentStatus)
                      //  Console.WriteLine("sds");
                    //Get the size of presentation
                    PresentationML.SlideSize slideInfo = presentation.SlideSize;
                    _presentationSizeX = slideInfo.Cx.Value;
                    _presentationSizeY = slideInfo.Cy.Value;

                    //Get the slidemaster scene object, background etc
                    PresentationML.SlideMaster slideMaster = _presentationDocument.PresentationPart.SlideMasterParts.ElementAt(0).SlideMaster; ;
                    currentMaster = slideMaster;
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
                        currentSlide = slidePart.Slide;
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

                        currentSlide = null;

                        //Add scene to presentation
                        _presentationObject.addScene(scene);
                        sceneCounter++;
                    }


                    _presentationObject.ConvertToYoobaUnits(_presentationSizeX, _presentationSizeY);
                }
          //  }
            /*catch(Exception e)
            {
                Console.WriteLine("Error reading '" + _path + "'");
                Console.WriteLine("\nError code: \n\n" + e);
            }*/
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

            List<SceneObject> imageObjectList = getImages(picture);
            foreach (SceneObject obj in imageObjectList)
                sceneObjectList.Add(obj);

            return sceneObjectList;
        }

        private List<SceneObject> getSceneObjects(PresentationML.GraphicFrame graphicFrame)
        {
            List<SceneObject> sceneObjectList = new List<SceneObject>();

            int boundsX = (int)graphicFrame.Transform.Offset.X,
                boundsY = (int)graphicFrame.Transform.Offset.Y;

            foreach (DrawingML.Table table in graphicFrame.Graphic.GraphicData.Descendants<DrawingML.Table>())
            {
                //Get the table style from the tabelStyles.xml (stored in tableStyle)
                DrawingML.TableStyleList tableStyleList = _presentationDocument.PresentationPart.TableStylesPart.TableStyleList;
                DrawingML.TableStyleEntry tableStyle = new DrawingML.TableStyleEntry();
                foreach (DrawingML.TableStyleId tableStyleId in table.TableProperties.Descendants<DrawingML.TableStyleId>())
                {
                    string tableId = tableStyleId.InnerText;
                    foreach (DrawingML.TableStyleEntry style in tableStyleList.Descendants<DrawingML.TableStyleEntry>())
                        if (style.StyleId == tableId)
                            tableStyle = style;
                }

                List<TableStyle> tableStyles = new List<TableStyle>();

                foreach(var style in tableStyle)
                {
                    TableStyle t = new TableStyle();
                    t.Type = style.LocalName;
                    t.FontSize = 1400;

                    foreach (var child in style)
                    {
                        if (child.LocalName == "tcTxStyle")
                        {
                            DrawingML.TableCellTextStyle tcTxStyle = (DrawingML.TableCellTextStyle)child;

                            if (tcTxStyle.Italic != null)
                                if (tcTxStyle.Italic.Value.ToString().ToLower() == "on" || tcTxStyle.Italic.Value.ToString().ToLower() == "true" ||
                                    tcTxStyle.Italic.Value.ToString().ToLower() == "t"  || tcTxStyle.Italic.Value.ToString().ToLower() == "1")
                                    t.Italic = true;

                            if (tcTxStyle.Bold != null)
                                if (tcTxStyle.Bold.Value.ToString().ToLower() == "on"   ||tcTxStyle.Bold.Value.ToString().ToLower() == "true" ||
                                    tcTxStyle.Bold.Value.ToString().ToLower() == "t"    || tcTxStyle.Bold.Value.ToString().ToLower() == "1"    )
                                    t.Bold = true;

                            string color = getColor(tcTxStyle);
                            if (color != "")
                                t.FontColor = color;

                            //FIX FONT AND FONT SIZE

                            //TODO
                        }
                        
                        if (child.LocalName == "tcStyle")
                        {
                            foreach (var tcStyleChild in child)
                            {
                                if (tcStyleChild.LocalName == "fill")
                                {
                                    DrawingML.FillProperties fill = (DrawingML.FillProperties)tcStyleChild;

                                    if (fill.SolidFill != null)
                                    {
                                        t.FillColor = getColor(fill.SolidFill);
                                        t.FillAlpha = getAlpha(fill.SolidFill);
                                    }
                                }

                                if (tcStyleChild.LocalName == "tcBdr")
                                {
                                    if (tcStyleChild.FirstChild != null)
                                    {
                                        if (tcStyleChild.FirstChild.FirstChild != null)
                                        {
                                            DrawingML.Outline ln = (DrawingML.Outline)tcStyleChild.FirstChild.FirstChild;

                                            if (ln.Width != null)
                                            {
                                                t.LineSize = ln.Width.Value;

                                                foreach (var lnChild in ln)
                                                {
                                                    if (lnChild.LocalName == "solidFill")
                                                    {
                                                        t.LineColor = getColor(lnChild);
                                                    }
                                                }
                                            }

                                        }
                                    }
                                }
                            }

                        }
                    }

                    tableStyles.Add(t);
                }

                int lastColIndex = -1, lastRowIndex = -1;

                //Store the column widths
                List<int> gridColWidthList = new List<int>();
                foreach (DrawingML.GridColumn gridCol in table.TableGrid)
                {
                    gridColWidthList.Add((int)gridCol.Width);
                    lastColIndex++;
                }

                foreach (DrawingML.TableRow tr in table.Descendants<DrawingML.TableRow>())
                    lastRowIndex++;

                int width, height, colIndex, rowIndex = 0, totalHeight = 0, totalWidth = 0;
                foreach (DrawingML.TableRow tr in table.Descendants<DrawingML.TableRow>())
                {
                    colIndex = 0;
                    foreach (DrawingML.TableCell tc in tr.Descendants<DrawingML.TableCell>())
                    {
                        width = gridColWidthList[colIndex];
                        height = (int)tr.Height;

                        if (tc.RowSpan != null)
                            height *= tc.RowSpan.Value;
                        
                        if(tc.GridSpan != null)
                            width *= tc.GridSpan.Value;
                        
                        SimpleSceneObject simpleSceneObjectShape = new SimpleSceneObject();
                        simpleSceneObjectShape.BoundsX = boundsX + totalWidth;
                        simpleSceneObjectShape.BoundsY = boundsY + totalHeight;
                        simpleSceneObjectShape.ClipWidth = width;
                        simpleSceneObjectShape.ClipHeight = height;

                        SimpleSceneObject simpleSceneObjectText = new SimpleSceneObject();
                        simpleSceneObjectText.BoundsX = boundsX + totalWidth;
                        simpleSceneObjectText.BoundsY = boundsY + totalHeight;
                        simpleSceneObjectText.ClipWidth = width;
                        simpleSceneObjectText.ClipHeight = height;

                        ShapeObject shapeObject = new ShapeObject(simpleSceneObjectShape, ShapeObject.shape_type.Rectangle);
                        TextObject textObject = new TextObject(simpleSceneObjectText);


                        //Set to default values
                        TableStyle defaultTableStyle = getTableStyle(tableStyles, "wholeTbl");
                        shapeObject.FillAlpha = defaultTableStyle.FillAlpha;
                        shapeObject.FillColor = defaultTableStyle.FillColor;
                        shapeObject.LineSize = defaultTableStyle.LineSize;
                        shapeObject.LineColor = defaultTableStyle.LineColor;
                        textObject.Color = defaultTableStyle.FontColor;
                        textObject.Size = defaultTableStyle.FontSize;

                        //if band column
                        if (table.TableProperties.BandRow != null)
                        {
                            if (rowIndex % 2 == 0)
                            {
                                shapeObject.setAttributes(getTableStyle(tableStyles, "band1H"));
                                textObject.setAttributes(getTableStyle(tableStyles, "band1H"));
                            }
                            else
                            {
                                shapeObject.setAttributes(getTableStyle(tableStyles, "band2H"));
                                textObject.setAttributes(getTableStyle(tableStyles, "band2H"));
                            } 
                        }

                        //if band column
                        if (table.TableProperties.BandColumn != null)
                        {
                            if (colIndex % 2 == 0)
                            {
                                shapeObject.setAttributes(getTableStyle(tableStyles, "band1V"));
                                textObject.setAttributes(getTableStyle(tableStyles, "band1V"));
                            }
                            else
                            {
                                shapeObject.setAttributes(getTableStyle(tableStyles, "band2V"));
                                textObject.setAttributes(getTableStyle(tableStyles, "band2V"));
                            }
                        }

                        //if first row
                        if (table.TableProperties.FirstRow != null && (rowIndex == 0))
                        {
                            shapeObject.setAttributes(getTableStyle(tableStyles, "firstRow"));
                            textObject.setAttributes(getTableStyle(tableStyles, "firstRow"));
                        }

                        //if first column
                        if (table.TableProperties.FirstColumn != null && (colIndex == 0))
                        {
                            shapeObject.setAttributes(getTableStyle(tableStyles, "firstCol"));
                            textObject.setAttributes(getTableStyle(tableStyles, "firstCol"));
                        }

                        //if last row
                        if (table.TableProperties.LastRow != null && (rowIndex == lastRowIndex))
                        {
                            shapeObject.setAttributes(getTableStyle(tableStyles, "lastRow"));
                            textObject.setAttributes(getTableStyle(tableStyles, "lastRow"));
                        }

                        //if last column
                        if (table.TableProperties.LastColumn != null && (colIndex == lastColIndex))
                        {
                            shapeObject.setAttributes(getTableStyle(tableStyles, "lastCol"));
                            textObject.setAttributes(getTableStyle(tableStyles, "lastCol"));
                        }

                        if (tc.TextBody != null)
                        {
                            int paragraphIndex = 0;
                            foreach (DrawingML.Paragraph p in tc.TextBody.Descendants<DrawingML.Paragraph>())
                            {

                                bool HasRun = false;
                                foreach (DrawingML.Run r in p.Descendants<DrawingML.Run>())
                                {
                                    HasRun = true;

                                    TextStyle textStyle = new TextStyle();
                                    TextFragment textFragment = new TextFragment();

                                    textFragment.Text = r.InnerText;
                                    textStyle.FontSize = textObject.Size;
                                    textStyle.FontColor = textObject.Color;
                                    textStyle.Bold = textObject.Bold;
                                    textStyle.Underline = textObject.Underline;
                                    textStyle.Italic = textObject.Italic;

                                    if (paragraphIndex != 0)
                                        textFragment.NewParagraph = true;

                                    //Override with specific run properties
                                    if (r.RunProperties != null)
                                    {
                                        DrawingML.RunProperties rPr = r.RunProperties;
                                        textStyle.Bold = (rPr.Bold != null) ? rPr.Bold.Value : textStyle.Bold;
                                        textStyle.Italic = (rPr.Italic != null) ? rPr.Italic.Value : textStyle.Italic;
                                        textStyle.FontSize = (rPr.FontSize != null) ? rPr.FontSize.Value : textStyle.FontSize;
                                        if (rPr.Underline != null)
                                            if (rPr.Underline.Value.ToString() == "sng")
                                                textStyle.Underline = true;

                                        //Get font color
                                        if (rPr.HasChildren)
                                            foreach (var rPrChild in rPr.ChildElements)
                                                if (rPrChild.LocalName == "solidFill")
                                                    textStyle.FontColor = getColor(rPrChild);
                                    }

                                    textObject.StyleList.Add(textStyle);
                                    textFragment.StyleId = textObject.StyleList.IndexOf(textStyle);
                                    textObject.FragmentsList.Add(textFragment);
                                }

                                if (!HasRun)
                                    if (textObject.FragmentsList.Count > 0)
                                        textObject.FragmentsList.Last().Breaks++;

                                paragraphIndex++;
                            }
                        }

                        if (tc.HorizontalMerge == null && tc.VerticalMerge == null)
                        {
                            sceneObjectList.Add(shapeObject);
                            sceneObjectList.Add(textObject);
                        }

                        colIndex++;
                        totalWidth += width;
                    }

                    rowIndex++;
                    totalHeight += (int)tr.Height;
                    totalWidth = 0;
                }
            }

            return sceneObjectList;
        }

        private TableStyle getTableStyle(List<TableStyle> list, string type)
        {
            foreach (TableStyle t in list)
                if (t.Type == type)
                    return t;

            return null;
        }

        private List<SceneObject> getSceneObjects(PresentationML.Shape shape)
        {
            List<SceneObject> sceneObjectList = new List<SceneObject>();

            int GchildOffX = 0, GchildOffY = 0, GchildExtX = 0, GchildExtY = 0, GoffX = 0, GoffY = 0, GextX = 0, GextY = 0, Grot = 0;

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

            bool isGroup = false;
            //Check for potential group shapes.
            IEnumerator<OpenXmlElement> parent = shape.Parent.GetEnumerator();
            while (parent.MoveNext())
            {
                if (parent.Current.LocalName == "grpSpPr")
                {

                    PresentationML.GroupShapeProperties grpSpPr = (PresentationML.GroupShapeProperties)parent.Current;

                    int rot = (grpSpPr.TransformGroup.Rotation != null) ? (int)grpSpPr.TransformGroup.Rotation : 0;

                    int offX = (int)grpSpPr.TransformGroup.Offset.X;
                    int offY = (int)grpSpPr.TransformGroup.Offset.Y;
                    int extX = (int)grpSpPr.TransformGroup.Extents.Cx;
                    int extY = (int)grpSpPr.TransformGroup.Extents.Cy;

                    int chOffX = (int)grpSpPr.TransformGroup.ChildOffset.X;
                    int chOffY = (int)grpSpPr.TransformGroup.ChildOffset.Y;
                    int chExtX = (int)grpSpPr.TransformGroup.ChildExtents.Cx;
                    int chExtY = (int)grpSpPr.TransformGroup.ChildExtents.Cy;

                    if ((rot != 0) || (offX != 0) || (offY != 0) || (extX != 0) || (extY != 0)
                        || (chOffX != 0) || (chOffY != 0) || (chExtX != 0) || (chExtY != 0))
                    {
                        GchildOffX = chOffX;
                        GchildOffY = chOffY;
                        GchildExtX = chExtX;
                        GchildExtY = chExtY;

                        GoffX = offX;
                        GoffY = offY;
                        GextX = extX;
                        GextY = extY;

                        Grot = rot;

                        shapeSimpleSceneObject.Rotation = rot;
                        shapeSimpleSceneObject.BoundsX = offX;
                        shapeSimpleSceneObject.BoundsY = offY;
                        shapeSimpleSceneObject.ClipWidth = extX;
                        shapeSimpleSceneObject.ClipHeight = extY;

                        isGroup = true;
                    }

                    break;

                }
            }

            bool HasBg = false, HasLine = false, HasValidGeometry = false, HasText = false, HasTransform = false,
                 solidFillColorChange = false, lineColorChange = false, gradientColorChange = false, IsLine = false;


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
                        if (isGroup)
                        {
                            double scaleX = (double)GextX / GchildExtX;
                            double scaleY = (double)GextY / GchildExtY;

                            if (spPr.Transform2D.Rotation != null)
                                shapeSimpleSceneObject.Rotation += (int)spPr.Transform2D.Rotation.Value;
                            
                            shapeSimpleSceneObject.BoundsX = (GoffX - GchildOffX + (int)spPr.Transform2D.Offset.X);
                            shapeSimpleSceneObject.BoundsY = (GoffY - GchildOffY + (int)spPr.Transform2D.Offset.Y);
                            shapeSimpleSceneObject.BoundsX = (int)((shapeSimpleSceneObject.BoundsX - GoffX) * scaleX) + GoffX;
                            shapeSimpleSceneObject.BoundsY = (int)((shapeSimpleSceneObject.BoundsY - GoffY) * scaleY) + GoffY;

                            shapeSimpleSceneObject.ClipWidth = (spPr.Transform2D.Extents.Cx != null) ? (int)(spPr.Transform2D.Extents.Cx * scaleX) : shapeSimpleSceneObject.ClipWidth;
                            shapeSimpleSceneObject.ClipHeight = (spPr.Transform2D.Extents.Cy != null) ? (int)(spPr.Transform2D.Extents.Cy * scaleY) : shapeSimpleSceneObject.ClipHeight;
                        
                            if (Grot != 0)
                            {
                                //Calculate the center of mass for the group and the child
                                double G_COM_X = GoffX + GextX / 2,
                                       G_COM_Y = GoffY + GextY / 2,
                                       C_COM_X = shapeSimpleSceneObject.BoundsX + shapeSimpleSceneObject.ClipWidth / 2,
                                       C_COM_Y = shapeSimpleSceneObject.BoundsY + shapeSimpleSceneObject.ClipHeight / 2;

                                //Calculate the distance between the two points
                                double distance = Math.Sqrt(Math.Pow(G_COM_X - C_COM_X, 2) + Math.Pow(G_COM_Y - C_COM_Y, 2));

                                //Calculate the angle in radians
                                double angleInRadians = (Grot / 60000) * Math.PI / 180;

                                //Calculate the diffrences in x and y between the centre of masses
                                double dy = Math.Sin(angleInRadians) * distance,
                                       dx = distance - (Math.Cos(angleInRadians) * distance);

                                //Handle the for different cases, the 4 quadrants
                                if ((C_COM_X <= G_COM_X) && (C_COM_X <= G_COM_X))
                                {
                                    shapeSimpleSceneObject.BoundsX += (int)Math.Round(dx);
                                    shapeSimpleSceneObject.BoundsY -= (int)Math.Round(dy);
                                }
                                else if ((C_COM_X <= G_COM_X) && (C_COM_X >= G_COM_X))
                                {
                                    shapeSimpleSceneObject.BoundsX += (int)Math.Round(dx);
                                    shapeSimpleSceneObject.BoundsY += (int)Math.Round(dy);
                                }
                                else if ((C_COM_X >= G_COM_X) && (C_COM_X <= G_COM_X))
                                {
                                    shapeSimpleSceneObject.BoundsX -= (int)Math.Round(dx);
                                    shapeSimpleSceneObject.BoundsY -= (int)Math.Round(dy);
                                }
                                else if ((C_COM_X >= G_COM_X) && (C_COM_X >= G_COM_X))
                                {
                                    shapeSimpleSceneObject.BoundsX -= (int)Math.Round(dx);
                                    shapeSimpleSceneObject.BoundsY += (int)Math.Round(dy);
                                }
                            }
                        
                        }
                        else
                        {
                            shapeSimpleSceneObject.Rotation = (spPr.Transform2D.Rotation != null) ? (int)spPr.Transform2D.Rotation.Value : shapeSimpleSceneObject.Rotation;
                            shapeSimpleSceneObject.BoundsX = (spPr.Transform2D.Offset.X != null) ? (int)spPr.Transform2D.Offset.X : shapeSimpleSceneObject.BoundsX;
                            shapeSimpleSceneObject.BoundsY = (spPr.Transform2D.Offset.Y != null) ? (int)spPr.Transform2D.Offset.Y : shapeSimpleSceneObject.BoundsY;
                            shapeSimpleSceneObject.ClipWidth = (spPr.Transform2D.Extents.Cx != null) ? (int)spPr.Transform2D.Extents.Cx : shapeSimpleSceneObject.ClipWidth;
                            shapeSimpleSceneObject.ClipHeight = (spPr.Transform2D.Extents.Cy != null) ? (int)spPr.Transform2D.Extents.Cy : shapeSimpleSceneObject.ClipHeight;
                        }

                        textSimpleSceneObject.Rotation = shapeSimpleSceneObject.Rotation;
                        textSimpleSceneObject.BoundsX = shapeSimpleSceneObject.BoundsX;
                        textSimpleSceneObject.BoundsY = shapeSimpleSceneObject.BoundsY;
                        textSimpleSceneObject.ClipWidth = shapeSimpleSceneObject.ClipWidth;
                        textSimpleSceneObject.ClipHeight = shapeSimpleSceneObject.ClipHeight;


                        powerPointText.Rotation = (int)Math.Round(shapeSimpleSceneObject.Rotation);
                        powerPointText.X = shapeSimpleSceneObject.BoundsX;
                        powerPointText.Y = shapeSimpleSceneObject.BoundsY;
                        powerPointText.Cx = shapeSimpleSceneObject.ClipWidth;
                        powerPointText.Cy = shapeSimpleSceneObject.ClipHeight;

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
                            if (prstGeom.Preset == "line")
                            {
                                IsLine = true;
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

                            shapeObject.FillAlpha = getAlpha(spPrChild);
                            shapeObject.FillType = "solid";

                        }   

                        if(spPrChild.LocalName == "gradFill")
                        {
                            shapeObject.GradientType = getGradientType((DrawingML.GradientFill)spPrChild);
                            shapeObject.GradientAngle = getGradientAngle((DrawingML.GradientFill)spPrChild);

                            HasBg = true;
                            List<string> colors = getColors(spPrChild);
                            List<int> alphas = getAlphas(spPrChild);

                            shapeObject.FillAlpha1 = alphas.First();
                            shapeObject.FillAlpha2 = alphas.Last();
                            shapeObject.FillColor1 = colors[0];
                            shapeObject.FillColor2 = colors[colors.Count-1];
                            shapeObject.FillType = "gradient";
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
                            if (IsLine)
                            {

                                if (shapeSimpleSceneObject.ClipWidth == 0)
                                    shapeSimpleSceneObject.ClipWidth = ln.Width.Value*100;
                                else if (shapeSimpleSceneObject.ClipHeight == 0)
                                    shapeSimpleSceneObject.ClipHeight = ln.Width.Value*100;

                                shapeObject = new ShapeObject(shapeSimpleSceneObject, ShapeObject.shape_type.Rectangle);

                            }
                            else
                            {
                                shapeObject.LineSize = (ln.Width != null) ? ln.Width.Value : shapeObject.LineSize;
                            }
                                
                            foreach (var lnChild in ln)
                            {
                                if (lnChild.LocalName == "solidFill")
                                {
                                    HasLine = true;
                                    lineColorChange = true;
                                    if (IsLine)
                                        shapeObject.FillColor = getColor(lnChild);
                                    else
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
                            for (int i = 0; i < 9; i++)
                            {
                                powerPointLevelList[i].Idx = powerPointText.Idx;
                                powerPointLevelList[i].Type = powerPointText.Type;
                                powerPointLevelList[i].Level = i;

                                PowerPointText layout = getPowerPointObject(_placeHolderListLayout, powerPointLevelList[i]);
                                if (layout != null)
                                {
                                    powerPointLevelList[i].setVisualAttribues(layout);
                                    powerPointLevelList[i].Level = i;
                                }
                            }

                            powerPointText = powerPointLevelList[0];
                        }

                    }

                    //Get the transform properties
                    if (HasTransform)
                    {

                        for (int i = 0; i < 9; i++)
                        {
                            powerPointLevelList[i].X = textSimpleSceneObject.BoundsX;
                            powerPointLevelList[i].Y = textSimpleSceneObject.BoundsY;
                            powerPointLevelList[i].Cx = textSimpleSceneObject.ClipWidth;
                            powerPointLevelList[i].Cy = textSimpleSceneObject.ClipHeight;
                        }
                    }

                    //if textbody contains listinfo, get that information
                    if (txBody.ListStyle != null)
                    {
                        listStyleList = getListStyles(txBody.ListStyle);
                        for (int i = 0; i < 9; i++)
                            powerPointLevelList[i].setVisualAttribues(listStyleList[i]);                   
                    }

                    if (!isGroup)
                    {
                        textSimpleSceneObject.BoundsX = powerPointText.X;
                        textSimpleSceneObject.BoundsY = powerPointText.Y;
                        textSimpleSceneObject.ClipWidth = powerPointText.Cx;
                        textSimpleSceneObject.ClipHeight = powerPointText.Cy;
                    }

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
                                if (temp != null)
                                {
                                    paragraghPPT.setVisualAttribues(temp);
                                }
                            }
                            else
                            {
                                paragraghPPT.Level = 0;
                            }
                        }

                        textObject.Align = (paragraghPPT.Alignment != "") ? paragraghPPT.Alignment : textObject.Align;
                        textObject.Color = (paragraghPPT.FontColor != "") ? paragraghPPT.FontColor : textObject.Color;
                        textObject.Size  = (paragraghPPT.FontSize > 0)    ? paragraghPPT.FontSize  : textObject.Size;

                        bool HasRun = false;
                        //Get the run properties
                        int runIndex = 0;
                        foreach (var pChild in p)
                        {
                            if (pChild.LocalName == "r")
                            {
                                DrawingML.Run r = (DrawingML.Run)pChild;

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
                                    if (rPr.Underline != null)
                                    {
                                        if (rPr.Underline.Value.ToString() == "sng")
                                        {
                                            runPPT.Underline = true;
                                        }
                                    }

                                    //Get font color
                                    if (rPr.HasChildren)
                                        foreach (var rPrChild in rPr.ChildElements)
                                            if (rPrChild.LocalName == "solidFill")
                                                runPPT.FontColor = getColor(rPrChild);

                                }

                                textStyle.Alignment = (runPPT.Alignment != "") ? runPPT.Alignment : textStyle.Alignment;
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

                            if (pChild.LocalName == "br")
                            {
                                TextFragment textFragment = new TextFragment();

                                if (textObject.FragmentsList.Count > 0)
                                {
                                    textFragment.StyleId = textObject.FragmentsList.Last().StyleId;
                                    textFragment.NewParagraph = true;
                                    textFragment.Text = "";
                                }
                                else
                                {
                                    TextStyle textStyle = new TextStyle();
                                    textObject.StyleList.Add(textStyle);
                                    textFragment.StyleId = textObject.StyleList.IndexOf(textStyle);
                                }

                                textObject.FragmentsList.Add(textFragment);
                                HasRun = true;
                            }
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
                                DrawingML.Outline line = (DrawingML.Outline)ln;

                                if (lineRefIndex == lineStyleIndex)
                                {

                                    if (line.Width != null)
                                        shapeObject.LineSize = line.Width.Value;


                                    foreach (var lineStyle in ln)
                                    {
                                        if (lineStyle.LocalName == "solidFill")
                                        {
                                            if (!lineColorChange)
                                            {
                                                HasLine = true;
                                                shapeObject.LineColor = getColor(lineStyle, color);
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
                                        if (!solidFillColorChange && !gradientColorChange)
                                        {
                                            HasBg = true;
                                            shapeObject.FillColor = getColor(fillStyle, color);
                                            shapeObject.FillAlpha = getAlpha(fillStyle);
                                            shapeObject.FillType = "solid";

                                        }
                                    }

                                    if (fillStyle.LocalName == "gradFill")
                                    {
                                        if (!gradientColorChange && !solidFillColorChange)
                                        {
                                            shapeObject.GradientType = getGradientType((DrawingML.GradientFill)fillStyle);
                                            shapeObject.GradientAngle = getGradientAngle((DrawingML.GradientFill)fillStyle);

                                            List<int> alphas = getAlphas(fillStyle);
                                            shapeObject.FillAlpha1 = alphas.First();
                                            shapeObject.FillAlpha2 = alphas.Last();

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
                sceneObjectList.Add(textObject);
                
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
                        {
                            shape.FillColor = getColor(bgStyle, color);
                            shape.FillAlpha = getAlpha(bgStyle);
                            sceneObjectList.Add(shape);
                        }
                            

                        if (bgStyle.LocalName == "gradFill")
                        {
                            shape.GradientType = getGradientType((DrawingML.GradientFill)bgStyle);
                            shape.GradientAngle = getGradientAngle((DrawingML.GradientFill)bgStyle);

                            List<string> colors = getColors(bgStyle, color);
                            List<int> gradientAlphaList = getAlphas(bgStyle);

                            shape.FillAlpha1 = gradientAlphaList.First();
                            shape.FillAlpha2 = gradientAlphaList.Last();
                            shape.FillColor1 = colors.First();
                            shape.FillColor2 = colors.Last();
                            shape.FillType = "gradient";
                            sceneObjectList.Add(shape);

                        }
                        if (bgStyle.LocalName == "blipFill")
                        {
                            List<SceneObject> images = getImages(bgStyle);
                            foreach(SceneObject obj in images)
                                sceneObjectList.Add(obj);
                        }
                    }
                    counter++;
                }
                

            }
            else if (background.FirstChild.GetType() == typeof(PresentationML.BackgroundProperties))
            {
                foreach (var bgType in background.FirstChild)
                {
                    if (bgType.LocalName == "solidFill")
                    {
                        foreach (DrawingML.SolidFill solidFill in background.Descendants<DrawingML.SolidFill>())
                        {
                            shape.FillType = "solid";
                            shape.FillColor = getColor(solidFill);
                            shape.FillAlpha = getAlpha(solidFill);
                        }
                        sceneObjectList.Add(shape);
                    }
                    if (bgType.LocalName == "gradFill")
                    {
                        foreach (DrawingML.GradientFill gradientFill in background.Descendants<DrawingML.GradientFill>())
                        {

                            shape.FillType = "gradient";
                            shape.GradientType = getGradientType(gradientFill);
                            shape.GradientAngle = getGradientAngle(gradientFill);

                            List<string> gradientColorList = getColors(gradientFill);
                            List<int> gradientAlphaList = getAlphas(gradientFill);

                            shape.FillAlpha1 = gradientAlphaList.First();
                            shape.FillAlpha2 = gradientAlphaList.Last();
                            shape.FillColor1 = gradientColorList.First();
                            shape.FillColor2 = gradientColorList.Last();
                        }
                        sceneObjectList.Add(shape);
                    }
                    if (bgType.LocalName == "blipFill")
                    {
                        List<SceneObject> images = getImages(background.FirstChild);
                        foreach (SceneObject obj in images)
                            sceneObjectList.Add(obj);
                    }
                }

            }

            return sceneObjectList;
        }

        private List<SceneObject> getImages(OpenXmlElement xmlElement)
        {
            List<SceneObject> tempImageList = new List<SceneObject>();
            SimpleSceneObject imageScenObject = new SimpleSceneObject();
            MediaFile MF = new MediaFile();
            if (xmlElement.GetType() == typeof(PresentationML.BackgroundProperties))
            {
                imageScenObject.ClipWidth = _presentationSizeX;
                imageScenObject.ClipHeight = _presentationSizeY;

                
                foreach (var child in xmlElement)
                {
                
                    if (child.LocalName == "blipFill")
                    {
                        DrawingML.BlipFill blipFill = (DrawingML.BlipFill)child;

                        string file_id = blipFill.Blip.Embed.Value;
                        foreach (var blipChild in blipFill.Blip)
                        {
                            if (blipChild.LocalName == "alphaModFix")
                            {
                                DrawingML.AlphaModulationFixed alphaMod = (DrawingML.AlphaModulationFixed)blipChild;
                                MF.Alpha = alphaMod.Amount;
                            }
                        }

                        if (_masterLevel)
                        {
                            
                            MF.ImagePart = (ImagePart)currentMaster.SlideMasterPart.GetPartById(file_id);
                            MF.ImageName = splitUriToImageName(MF.ImagePart.Uri);
                        }
                        if (_slideLevel)
                        {

                            MF.ImagePart = (ImagePart)currentSlide.SlidePart.GetPartById(file_id);
                            MF.ImageName = splitUriToImageName(MF.ImagePart.Uri);
                        }

                    }

                }
            }
            else if (xmlElement.GetType() == typeof(PresentationML.Picture))
            {
                foreach (var child in xmlElement)
                {

                    if (child.LocalName == "nvPicPr")
                    {
                        PresentationML.NonVisualPictureProperties nvPicPr = (PresentationML.NonVisualPictureProperties)child;

                    }
                    if (child.LocalName == "blipFill")
                    {
                        PresentationML.BlipFill blipFill = (PresentationML.BlipFill)child;

                        string file_id = blipFill.Blip.Embed.Value;

                        foreach (var blipChild in blipFill.Blip)
                        {
                            if (blipChild.LocalName == "alphaModFix")
                            {
                                DrawingML.AlphaModulationFixed alphaMod = (DrawingML.AlphaModulationFixed)blipChild;
                                MF.Alpha = alphaMod.Amount;
                            }
                        }

                        if (_masterLevel)
                        {

                            MF.ImagePart = (ImagePart)currentMaster.SlideMasterPart.GetPartById(file_id);
                            MF.ImageName = splitUriToImageName(MF.ImagePart.Uri);
                        }
                        if (_slideLevel)
                        {

                            MF.ImagePart = (ImagePart)currentSlide.SlidePart.GetPartById(file_id);
                            MF.ImageName = splitUriToImageName(MF.ImagePart.Uri);
                        }

                    }
                    if (child.LocalName == "spPr")
                    {
                        PresentationML.ShapeProperties spPr = (PresentationML.ShapeProperties)child;

                        
                        imageScenObject.BoundsX = (spPr.Transform2D.Offset.X != null) ? (int)spPr.Transform2D.Offset.X : imageScenObject.BoundsX;
                        imageScenObject.BoundsY = (spPr.Transform2D.Offset.Y != null) ? (int)spPr.Transform2D.Offset.Y : imageScenObject.BoundsY;
                        imageScenObject.ClipWidth = (spPr.Transform2D.Extents.Cx != null) ? (int)spPr.Transform2D.Extents.Cx : imageScenObject.ClipWidth;
                        imageScenObject.ClipHeight = (spPr.Transform2D.Extents.Cy != null) ? (int)spPr.Transform2D.Extents.Cy : imageScenObject.ClipHeight;

                        MF.OffX = imageScenObject.BoundsX;
                        MF.OffX = imageScenObject.BoundsY;
                        MF.ExtX = imageScenObject.ClipWidth;
                        MF.ExtY = imageScenObject.ClipHeight;
                    }
                }

            }

            //Console.WriteLine(MF.toString());
            

            //if connection and response != null
            //read information needed clipId etc and apply to imageScenObject.
            foreach (ImageResponse ir in _listOfimageResponse)
            {
                /*if (ir.Label == MF.ImageName)
                    imageScenObject.ClipID = ir.ClipID;*/
                
                if (ir.Label == "salesman.jpg")
                    imageScenObject.ClipID = ir.ClipID;
            }

            ImageObject imageObj = new ImageObject(imageScenObject);
            if(imageObj != null)
                tempImageList.Add(imageObj);

            return tempImageList;
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

            foreach (var colorValue in xmlElement.Descendants<DrawingML.RgbColorModelHex>())
            {
                if (colorValue.Val.ToString() == "phClr")
                    color = phClr;
                else
                    color = colorValue.Val.ToString();

                if (colorValue.HasChildren)
                {
                    string transformedColor = colorTransforms(colorValue, color);
                    color = transformedColor;
                    return color;
                }
                else
                    return color;
            }

            foreach (var colorValue in xmlElement.Descendants<DrawingML.SchemeColor>())
            {
                if (colorValue.Val.ToString() == "phClr")
                    color = phClr;
                else
                    color = getColorFromTheme(colorValue.Val.ToString());

                if (colorValue.HasChildren)
                {
                    string transformedColor = colorTransforms(colorValue, color);
                    color = transformedColor;
                    return color;
                }
                else
                    return color;
            }

            return "";

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
                if (child.LocalName == "alpha")
                {
                    DrawingML.Alpha alpha = (DrawingML.Alpha)child;
                    return alpha.Val.Value;
                }

            //Default
            return 100000;
        }
        private static string splitUriToImageName(Uri uri)
        {

            string[] image = uri.OriginalString.Split('/');
            string[] image_name = image[3].Split('.');

            return image_name[0] + "." + image_name[1];
        }

        public List<int> getAlphas(OpenXmlElement xmlElement)
        {
            List<int> alphas = new List<int>();
            OpenXmlElement elementType = xmlElement.FirstChild;

            foreach (DrawingML.GradientStop gs in elementType.Descendants<DrawingML.GradientStop>())
                alphas.Add(getAlpha(gs));

            return alphas;
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

        public void getImagesFromDatabase()
        {
            string response = "<assets>" + 
              "<asset clipID='9457e300-517a-4fe8-9ebe-507c1be16bed'  itemType='clip' width='500' height='628' label='sven_sm.png' size='409165' uploadDate='2015-02-12 13:57' url='https://d1lskzugeos37d.cloudfront.net/f969a3ce-9cfb-41aa-8add-a2fa00e44eda/2a6/2a62fa1b-d6f1-44ea-a85e-13b9800afff3.png'/>" +
              "<asset clipID='ff257113-dda6-404a-b7de-873d72c12836' itemType='clip' width='291' height='510' label='dataSourceScene.png' size='15555' uploadDate='2014-12-20 23:17' url='https://d1lskzugeos37d.cloudfront.net/f969a3ce-9cfb-41aa-8add-a2fa00e44eda/008/008e160c-1807-4bb0-9180-905b50e2f322.png'/>" +
              "<asset clipID='5825ccd6-8477-481b-9d34-ba873c9af577' itemType='clip' width='500' height='59' label='button-down.png' size='3007' uploadDate='2014-11-12 11:06' url='https://d1lskzugeos37d.cloudfront.net/f969a3ce-9cfb-41aa-8add-a2fa00e44eda/7a0/7a04e656-2919-4915-9f9f-260ba398a6fa.png'/>" +
              "<asset clipID='4bbb885d-1b45-4840-984c-c5ed5d0b3483' itemType='clip' width='963' height='912' label='1.PNG' size='71686' uploadDate='2015-02-20 15:35' url='https://d1lskzugeos37d.cloudfront.net/f969a3ce-9cfb-41aa-8add-a2fa00e44eda/5a4/5a438b08-2a64-4899-970b-82df591faa77.png'/>" +
              "<asset clipID='7a769473-ad60-413e-a477-cd0c4d1ca328' itemType='clip' width='3264' height='2448' label='davids-snurra.jpg' size='417399' uploadDate='2014-11-28 11:13' url='https://d1lskzugeos37d.cloudfront.net/f969a3ce-9cfb-41aa-8add-a2fa00e44eda/903/90375161-11ba-478f-b738-d5aee2d12216.jpg'/>" +
              "<asset clipID='4938d5b0-0a24-4977-b218-f7612f6e4630' itemType='clip' width='200' height='200' label='salesman.jpg' size='22343' uploadDate='2014-11-05 13:38' url='https://d1lskzugeos37d.cloudfront.net/f969a3ce-9cfb-41aa-8add-a2fa00e44eda/a86/a864d9fd-4c2e-4b9f-a0f8-a3b4d1e6fd77.jpg'/>" +
              "</assets>";

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response);
            foreach (XmlElement asset in doc.FirstChild.ChildNodes)
            {
                if (asset.LocalName == "asset")
                {
                    ImageResponse IR = new ImageResponse();

                    foreach (XmlAttribute attr in asset.Attributes)
                    {
                        if (attr.LocalName == "clipID")
                            IR.ClipID = attr.Value;
                        if (attr.LocalName == "itemType")
                            IR.ItemType = attr.Value;
                        if (attr.LocalName == "width")
                            IR.Width = attr.Value;
                        if (attr.LocalName == "height")
                            IR.Height = attr.Value;
                        if (attr.LocalName == "label")
                            IR.Label = attr.Value;
                        if (attr.LocalName == "size")
                            IR.Size = attr.Value;
                        if (attr.LocalName == "uploadDate")
                            IR.UploadDate = attr.Value;
                        if (attr.LocalName == "url")
                            IR.Url = attr.Value;
                    }
                    
                    _listOfimageResponse.Add(IR);
                }

            }

            
        }
        
    }
}
