using System;
using System.Data.SqlClient;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;

namespace BookMarks2
{
    public class BookmarkOpenxml
    {
        // Fetch binary data (FIRST_SIGNATURE) from database by user_name
        public static byte[] GetSignatureFromDatabase(string connectionString, string userName)
        {
            byte[] signatureData = null;

            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = "SELECT [FIRST_SIGNATURE] FROM [CentralUserInfo].[dbo].[users] WHERE [user_name] = @UserName";
                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@UserName", userName);

                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            signatureData = reader["FIRST_SIGNATURE"] as byte[];
                        }
                    }
                }
            }

            return signatureData;
        }

        // Insert PNG into a Word document at a specified bookmark
        public static void InsertPngAtBookmark(string wordFilePath, string bookmarkName, byte[] pngData)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, true))
            {
                // Find the bookmark
                var bookmark = wordDoc.MainDocumentPart.Document.Descendants<BookmarkStart>()
                    .FirstOrDefault(b => b.Name == bookmarkName);

                if (bookmark == null)
                {
                    throw new Exception($"Bookmark '{bookmarkName}' not found in the document.");
                }

                // Create an image part and add the PNG data
                var mainPart = wordDoc.MainDocumentPart;
                var imagePart = mainPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png);

                using (var stream = new MemoryStream(pngData))
                {
                    imagePart.FeedData(stream);
                }

                // Get image dimensions in EMUs
                var (width, height) = GetImageDimensionsFromStream(new MemoryStream(pngData));

                // Create the image element
                var drawingElement = CreateImageElement(mainPart.GetIdOfPart(imagePart), width, height);
                var imageRun = new DocumentFormat.OpenXml.Wordprocessing.Run(drawingElement);

                // Insert the image after the bookmark
                bookmark.Parent.InsertAfter(imageRun, bookmark);
                wordDoc.MainDocumentPart.Document.Save();
            }
        }

        // Utility to get dimensions from PNG binary data
        private static (long width, long height) GetImageDimensionsFromStream(Stream imageStream)
        {
            using (var image = SixLabors.ImageSharp.Image.Load(imageStream))
            {
                const int emusPerInch = 914400;
                var dpiX = image.Metadata.HorizontalResolution;
                var dpiY = image.Metadata.VerticalResolution;
                return (
                    (long)(image.Width * emusPerInch / dpiX),
                    (long)(image.Height * emusPerInch / dpiY)
                );
            }
        }

        // Create image element
        private static Drawing CreateImageElement(string relationshipId, long width, long height)
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
                                        Name = "Image"
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
