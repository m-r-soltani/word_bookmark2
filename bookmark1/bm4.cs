using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace bm4
{

    class Program
    {
        public static void Main()
        {
            string filePath = @"C:\mydocs\f2.docx";
            string imagePath = @"C:\mydocs\img\1.png";

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                // Add image to the main document part
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (FileStream stream = new FileStream(imagePath, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                string imagePartId = mainPart.GetIdOfPart(imagePart);

                // Insert image at bookmarks in body, header, and footer
                InsertImageAtBookmark(mainPart.Document.Body, "BodyBookmark", imagePartId);

                foreach (var header in mainPart.HeaderParts)
                {
                    InsertImageAtBookmark(header.Header, "HeaderBookmark", imagePartId);
                }

                foreach (var footer in mainPart.FooterParts)
                {
                    InsertImageAtBookmark(footer.Footer, "FooterBookmark", imagePartId);
                }

                mainPart.Document.Save();
            }

            Console.WriteLine("Image inserted successfully.");
        }

        private static void InsertImageAtBookmark(OpenXmlElement container, string bookmarkName, string imagePartId)
        {
            BookmarkStart bookmarkStart = container.Descendants<BookmarkStart>()
                .FirstOrDefault(b => b.Name == bookmarkName);

            if (bookmarkStart != null)
            {
                DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run();
                Drawing drawing = CreateImageDrawing(imagePartId);
                run.Append(drawing);
                bookmarkStart.Parent.InsertAfterSelf(run);
            }
        }

        private static Drawing CreateImageDrawing(string relationshipId, long width = 120 * 914400 / 96, long height= 70 * 914400 / 96)
        {
            return new Drawing(
                new Inline(
                    new Extent { Cx = width, Cy = height },
                    new EffectExtent
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DocProperties
                    {
                        Id = (UInt32Value)1U,
                        Name = "Picture"
                    },
                    new DocumentFormat.OpenXml.Drawing.NonVisualGraphicFrameDrawingProperties(new GraphicFrameLocks { NoChangeAspect = true }),
                    new Graphic(
                        new GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
                                    {
                                        Id = (UInt32Value)0U,
                                        Name = "New Image.jpg"
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
                                ),
                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                    new Blip { Embed = relationshipId },
                                    new Stretch(new FillRectangle())
                                ),
                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                    new Transform2D(
                                        new Offset { X = 0L, Y = 0L },
                                        new Extents { Cx = width, Cy = height }
                                    ),
                                    new PresetGeometry(new AdjustValueList())
                                    { Preset = ShapeTypeValues.Rectangle }
                                )
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)0U,
                    DistanceFromRight = (UInt32Value)0U,
                    EditId = "50D07946"
                });
        }
    }

}





