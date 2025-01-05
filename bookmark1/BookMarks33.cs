using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Metadata;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Pictures;


namespace BookMarks3
{
    //public class BookmarkOpenxml
    //{
    //    public static void UpdateBookmarks(string filePath, Dictionary<string, string> bookmarksContent, string connectionString)
    //    {
    //        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
    //        {
    //            foreach (var entry in bookmarksContent)
    //            {
    //                string bookmarkName = entry.Key;
    //                string content = entry.Value;

    //                // Find the bookmark
    //                var bookmark = FindBookmark(wordDoc, bookmarkName);

    //                if (bookmark != null)
    //                {
    //                    // Locate the bookmark's text content
    //                    var nextSibling = bookmark.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Run>();

    //                    if (File.Exists(content)) // Check if the value is an image file path
    //                    {
    //                        nextSibling?.Remove();
    //                        InsertImageAtBookmark(wordDoc, bookmark, File.ReadAllBytes(content));
    //                    }
    //                    else if (IsBinaryContent(content)) // Handle binary data
    //                    {
    //                        byte[] imageData = GetImageFromDatabase(content, connectionString);
    //                        if (imageData != null)
    //                        {
    //                            nextSibling?.Remove();
    //                            InsertImageAtBookmark(wordDoc, bookmark, imageData);
    //                        }
    //                    }
    //                    else // Treat the value as text
    //                    {
    //                        if (nextSibling != null)
    //                        {
    //                            var textElement = nextSibling.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>();
    //                            if (textElement != null)
    //                            {
    //                                textElement.Text = content;
    //                            }
    //                        }
    //                        else
    //                        {
    //                            var newRun = new DocumentFormat.OpenXml.Wordprocessing.Run(
    //                                new DocumentFormat.OpenXml.Wordprocessing.Text(content));
    //                            bookmark.Parent.InsertAfter(newRun, bookmark);
    //                        }
    //                    }
    //                }
    //                else
    //                {
    //                    Console.WriteLine($"Bookmark '{bookmarkName}' not found in the document.");
    //                }
    //            }

    //            wordDoc.MainDocumentPart.Document.Save();
    //        }
    //    }

    //    private static void InsertImageAtBookmark(WordprocessingDocument wordDoc, DocumentFormat.OpenXml.Wordprocessing.BookmarkStart bookmark, byte[] imageData)
    //    {
    //        var mainPart = wordDoc.MainDocumentPart;
    //        var imagePart = mainPart.AddImagePart(ImagePartType.Png);

    //        using (var stream = new MemoryStream(imageData))
    //        {
    //            imagePart.FeedData(stream);
    //        }

    //        var (width, height) = GetImageDimensions(imageData);
    //        var drawingElement = CreateImageElement(mainPart.GetIdOfPart(imagePart), width, height);

    //        var imageRun = new DocumentFormat.OpenXml.Wordprocessing.Run(drawingElement);
    //        bookmark.Parent.InsertAfter(imageRun, bookmark);
    //    }

    //    private static DocumentFormat.OpenXml.Wordprocessing.Drawing CreateImageElement(string relationshipId, long width, long height)
    //    {
    //        return new DocumentFormat.OpenXml.Wordprocessing.Drawing(
    //            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
    //                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = width, Cy = height },
    //                new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent
    //                {
    //                    LeftEdge = 0L,
    //                    TopEdge = 0L,
    //                    RightEdge = 0L,
    //                    BottomEdge = 0L
    //                },
    //                new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties
    //                {
    //                    Id = (UInt32Value)1U,
    //                    Name = "Picture"
    //                },
    //                new DocumentFormat.OpenXml.Drawing.NonVisualGraphicFrameDrawingProperties(
    //                    new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks { NoChangeAspect = true }
    //                ),
    //                new DocumentFormat.OpenXml.Drawing.Graphic(
    //                    new DocumentFormat.OpenXml.Drawing.GraphicData(
    //                        new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
    //                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
    //                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
    //                                {
    //                                    Id = (UInt32Value)0U,
    //                                    Name = "New Image"
    //                                },
    //                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
    //                            ),
    //                            new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
    //                                new DocumentFormat.OpenXml.Drawing.Blip { Embed = relationshipId },
    //                                new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())
    //                            ),
    //                            new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
    //                                new DocumentFormat.OpenXml.Drawing.Transform2D(
    //                                    new DocumentFormat.OpenXml.Drawing.Offset { X = 0L, Y = 0L },
    //                                    new DocumentFormat.OpenXml.Drawing.Extents { Cx = width, Cy = height }
    //                                ),
    //                                new DocumentFormat.OpenXml.Drawing.PresetGeometry(new DocumentFormat.OpenXml.Drawing.AdjustValueList())
    //                                { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }
    //                            )
    //                        )
    //                    )
    //                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
    //                )
    //            )
    //            {
    //                DistanceFromTop = (UInt32Value)0U,
    //                DistanceFromBottom = (UInt32Value)0U,
    //                DistanceFromLeft = (UInt32Value)0U,
    //                DistanceFromRight = (UInt32Value)0U
    //            });
    //    }

    //    private static (long width, long height) GetImageDimensions(byte[] imageData)
    //    {
    //        using (var image = Image.Load(imageData))
    //        {
    //            const int emusPerInch = 914400;
    //            var dpiX = image.Metadata.HorizontalResolution;
    //            var dpiY = image.Metadata.VerticalResolution;

    //            return (
    //                (long)(image.Width * emusPerInch / dpiX),
    //                (long)(image.Height * emusPerInch / dpiY)
    //            );
    //        }
    //    }

    //    public static DocumentFormat.OpenXml.Wordprocessing.BookmarkStart FindBookmark(WordprocessingDocument wordDoc, string bookmarkName)
    //    {
    //        return wordDoc.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.BookmarkStart>()
    //            .FirstOrDefault(b => b.Name == bookmarkName);
    //    }

    //    private static byte[] GetImageFromDatabase(string userName, string connectionString)
    //    {
    //        using (var connection = new SqlConnection(connectionString))
    //        using (var command = new SqlCommand("SELECT FIRST_SIGNATURE FROM [CentralUserInfo].[dbo].[users] WHERE user_name = @userName", connection))
    //        {
    //            command.Parameters.AddWithValue("@userName", userName);
    //            connection.Open();
    //            var result = command.ExecuteScalar();
    //            return result != DBNull.Value ? (byte[])result : null;
    //        }
    //    }

    //    private static bool IsBinaryContent(string input)
    //    {
    //        return input.StartsWith("0x") || input.All(c => char.IsLetterOrDigit(c));
    //    }
    //}
}
