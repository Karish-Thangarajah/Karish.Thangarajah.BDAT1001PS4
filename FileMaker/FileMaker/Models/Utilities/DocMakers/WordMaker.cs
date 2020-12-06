using FileMaker.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;


namespace FileMaker.Models.Utilities.DocMakers
{
    class WordMaker
    {
        public static void CreateWordprocessingDocument(string filepath)
        {
            // Create a document by supplying the filepath. 
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
            }
        }

        public static void OpenAndAddText(string filepath, Student student)
        {
            // Open a WordprocessingDocument for editing using the filepath.
            WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(filepath, true);



            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            if (para.ParagraphProperties == null)
            {
                para.ParagraphProperties = new ParagraphProperties();
            }

            para.ParagraphProperties.PageBreakBefore = new PageBreakBefore();

            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text($"Hello, my name is {student.FirstName} {student.LastName}."));

            if (student.StudentCode == "200471940")
            {
                string relationId = AddGraph(wordprocessingDocument, @"C:\Users\karis\OneDrive\Desktop\BDAT1001\Assignment 4\myimage.jpg");
                InsertImage(wordprocessingDocument, run, @"C:\Users\karis\OneDrive\Desktop\BDAT1001\Assignment 4\myimage.jpg");

            }
            else
            {
                string relationId = AddGraph(wordprocessingDocument, @"C:\Users\karis\OneDrive\Desktop\BDAT1001\Assignment 4\unknown.jpg");
                InsertImage(wordprocessingDocument, run, @"C:\Users\karis\OneDrive\Desktop\BDAT1001\Assignment 4\unknown.jpg");
            }



            // Close the handle explicitly.
            wordprocessingDocument.Close();
        }

        private static string AddGraph(WordprocessingDocument wpd, string filepath)
        {
            ImagePart ip = wpd.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);
            using (FileStream fs = new FileStream(filepath, FileMode.Open))
            {
                if (fs.Length == 0) return string.Empty;
                ip.FeedData(fs);
            }

            return wpd.MainDocumentPart.GetIdOfPart(ip);
        }

        private static void InsertImage(WordprocessingDocument wpd, OpenXmlElement parent, string filepath)
        {
            string relationId = AddGraph(wpd, filepath);
            if (!string.IsNullOrEmpty(relationId))
            {
                var draw = new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = 2095500L, Cy = 2286000L },
                        new DW.EffectExtent()
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties()
                        {
                            Id = (UInt32Value)1U,
                            Name = "Picture 1"
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties()
                                        {
                                            Id = (UInt32Value)0U,
                                            Name = relationId
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                        new PIC.BlipFill(
                                            new A.Blip(
                                                new A.BlipExtensionList(
                                                    new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" })
                                                )
                                            {
                                                Embed = relationId,
                                                CompressionState =
                                                A.BlipCompressionValues.Print
                                            },
                                                new A.Stretch(
                                                    new A.FillRectangle())),
                                                    new PIC.ShapeProperties(
                                                        new A.Transform2D(
                                                            new A.Offset() { X = 0L, Y = 0L },
                                                            new A.Extents() { Cx = 2095500L, Cy = 2286000L }),
                                                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                                                            )
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U,
                        EditId = "50D07946"
                    });

                parent.Append(draw);
            }
        }

    }
}
