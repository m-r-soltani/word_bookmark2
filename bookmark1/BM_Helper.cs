using System.Drawing;
using System;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml;
using static System.Net.Mime.MediaTypeNames;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Metadata;
using System.Security.Cryptography.Xml;
using System.Xml.Linq;
using Microsoft.Extensions.Configuration;
using QRCoder;
using System.Linq;
using DocumentFormat.OpenXml.Office2016.Presentation.Command;
using System.Diagnostics;
using System.Text;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Drawing.Charts;


namespace BMH
{
    //BookMarkHelper
    public class BMH
    {
        ///////////////////////////////////////////Update Text Bookmarks///////////////////////////////////////////
        public static void UpdateTextBookmarks(string filePath, Dictionary<string, string> bookmarksContent)
        {
            bool inserted=false;
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                foreach (var entry in bookmarksContent)
                {
                    var bookmark= FindBookmark(wordDoc, entry.Key);
                    if (bookmark != null)
                    {
                        // Empty the content between the bookmarks
                        EmptyBookmark(wordDoc, bookmark.Value.bookmarkStart, bookmark.Value.bookmarkEnd);
                        // Insert the new content at the bookmark
                        inserted = InsertTextAtBookmark(wordDoc, bookmark.Value.bookmarkStart, entry.Value);
                    }
                    else {
                        Debug.WriteLine("Bookmark " + entry.Key + "Not Found");
                    }
                    
                }
                if (inserted)
                {
                    wordDoc.MainDocumentPart.Document.Save();
                }
                else {
                    //error ?
                    Console.WriteLine("Insertion Failed!");
                }
            }
        }

        public static bool InsertTextAtBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string newText)
        {
            try
            {
                // Find the parent run or create a new one
                var runElement = bookmarkStart.Parent.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().FirstOrDefault();
                if (runElement == null)
                {
                    runElement = new DocumentFormat.OpenXml.Wordprocessing.Run();
                    bookmarkStart.InsertAfterSelf(runElement);
                }

                // Clear existing text in the run
                var textElements = runElement.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>().ToList();
                foreach (var text in textElements)
                {
                    text.Remove();
                }

                // Split the new text by line breaks and add it to the run
                var lines = newText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                for (int i = 0; i < lines.Length; i++)
                {
                    if (!string.IsNullOrEmpty(lines[i]))
                    {
                        var textElement = new DocumentFormat.OpenXml.Wordprocessing.Text(lines[i])
                        {
                            Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve
                        };
                        runElement.AppendChild(textElement);
                    }

                    // Add a line break if it's not the last line
                    if (i < lines.Length - 1)
                    {
                        runElement.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Break());
                    }
                }

                Console.WriteLine($"Text successfully inserted at bookmark: {bookmarkStart.Name}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during text insertion: {ex.Message}");
                return false;
            }
        }

        ///////////////////////////////////////////Update Image Bookmarks///////////////////////////////////////////
        public static void UpdateImageBookmarks(string filePath, Dictionary<string, string> bookmarksContent)
        {
            bool inserted = false;
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                foreach (var entry in bookmarksContent)
                {
                    var bookmark = FindBookmark(wordDoc, entry.Key);
                    if (bookmark != null)
                    {
                        // Empty the content between the bookmarks
                        EmptyBookmark(wordDoc, bookmark.Value.bookmarkStart, bookmark.Value.bookmarkEnd);
                            if (File.Exists(entry.Value))
                            {
                                inserted = InsertImageAtBookmark(wordDoc, bookmark.Value.bookmarkStart, entry.Value);
                            }
                            else
                            {
                            //Console.OutputEncoding = Encoding.UTF8;
                            Console.WriteLine("File Not Found" + entry.Key + " " + entry.Value);
                            }
                    }
                    else
                    {
                        Debug.WriteLine("Bookmark " + entry.Key + "Not Found");
                    }

                }
                if (inserted)
                {
                    wordDoc.MainDocumentPart.Document.Save();
                }
                else
                {
                    //error ?
                    Console.WriteLine("Insertion Failed!!!!!!!!!!!!!!!!!" );
                }
            }
        }

        ///////////////////////////////////////////Update Binary Image Bookmarks///////////////////////////////////////////
        public static void UpdateBinaryImageBookmarks(string filePath, Dictionary<string, string> bookmarksContent)
        {
            bool inserted = false;
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                foreach (var entry in bookmarksContent)
                {
                    var bookmark = FindBookmark(wordDoc, entry.Key);
                    if (bookmark != null)
                    {
                        // Empty the content between the bookmarks
                        EmptyBookmark(wordDoc, bookmark.Value.bookmarkStart, bookmark.Value.bookmarkEnd);

                        // Fetch Binary Data By Username
                        byte[] imageData = GetImageDataFromDatabase(entry.Value);
                        if (imageData != null)
                        {
                            string tempFilePath = SaveImageToTempFile(imageData);
                            if (File.Exists(tempFilePath))
                            {
                                //insert image To File
                                inserted=InsertImageAtBookmark(wordDoc, bookmark.Value.bookmarkStart, tempFilePath);

                            }
                            else
                            {
                                Console.WriteLine("Temp File Not Found" + entry.Key + " " + entry.Value);
                            }
                        }
                        else
                        {
                            Console.WriteLine("tempFilePath Not Found!");
                        }
                    }
                    else
                    {
                        Debug.WriteLine("Bookmark " + entry.Key + "Not Found");
                    }

                }
                if (inserted)
                {
                    wordDoc.MainDocumentPart.Document.Save();
                }
                else
                {
                    //error ?
                    Console.WriteLine("Insertion Failed!");
                }
            }
        }

        public static byte[] GetImageDataFromDatabase(string userName)
        {
            string connectionString = "Server=localhost;Database=CenteralUserInfo;User Id=sa;Password=Aa@12345;";
            string query = "SELECT FIRST_SIGNATURE FROM CenteralUserInfo.dbo.Pargar_USER_SIGN WHERE USER_NAME = @UserName";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@UserName", userName);
                        connection.Open();

                        var result = command.ExecuteScalar();

                        if (result == null || result == DBNull.Value)
                        {
                            Console.WriteLine($"No data found for user: {userName}");
                            return null;
                        }

                        if (result is string hexString)
                        {
                            Console.WriteLine($"Hex string data retrieved for user: {userName}");

                            try
                            {
                                return ConvertHexStringToByteArray(hexString);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error converting hex string to byte array: {ex.Message}");
                                return null;
                            }
                        }

                        if (result is byte[] byteArray)
                        {
                            Console.WriteLine($"Binary data retrieved for user: {userName}");
                            return byteArray;
                        }

                        if (result is string base64String) // Check if the data is a string
                        {
                            Console.WriteLine($"Base64 string data retrieved for user: {userName}");
                            return Convert.FromBase64String(base64String); // Decode Base64 string to binary
                        }
                        else
                        {
                            Console.WriteLine($"Unexpected data type for user: {userName}. Data type: {result.GetType()}");
                            return null;
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving data for user: {userName}. Exception: {ex.Message}");
                return null;
            }
        }

        public static byte[] ConvertHexStringToByteArray(string hexString)
        {
            if (string.IsNullOrWhiteSpace(hexString))
                throw new ArgumentException("Input string cannot be null or empty.");

            // Remove the "0x" prefix if present
            if (hexString.StartsWith("0x"))
                hexString = hexString.Substring(2);

            // Ensure even length
            if (hexString.Length % 2 != 0)
                throw new ArgumentException("Hex string has an invalid length.");

            byte[] byteArray = new byte[hexString.Length / 2];
            for (int i = 0; i < hexString.Length; i += 2)
            {
                byteArray[i / 2] = Convert.ToByte(hexString.Substring(i, 2), 16);
            }

            return byteArray;
        }


        ///////////////////////////////////////////Update QrCode Bookmarks///////////////////////////////////////////
        public static void UpdateQrcodeBookmarks(string filePath, Dictionary<string, string> bookmarksContent)
        {
            bool inserted = false;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                foreach (var entry in bookmarksContent)
                {
                    var bookmark = FindBookmark(wordDoc, entry.Key);
                    if (bookmark != null)
                    {
                        // Clear existing bookmark content
                        EmptyBookmark(wordDoc, bookmark.Value.bookmarkStart, bookmark.Value.bookmarkEnd);

                        // Generate QR code image as a byte array
                        var qrCodeData = GenerateQRCodeAsByteArray(entry.Value);
                        if (qrCodeData != null)
                        {
                            // Save QR code image temporarily
                            string tempFilePath = SaveImageToTempFile(qrCodeData);
                            if (File.Exists(tempFilePath))
                            {
                                // Insert the image at the bookmark
                                inserted = InsertImageAtBookmark(wordDoc, bookmark.Value.bookmarkStart, tempFilePath);
                                File.Delete(tempFilePath); // Clean up the temporary file
                            }
                            else
                            {
                                Console.WriteLine("Failed to save the QR code image.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Failed to generate QR code data.");
                        }
                    }
                    else
                    {
                        Debug.WriteLine($"Bookmark '{entry.Key}' not found.");
                    }
                }

                // Save changes to the document
                if (inserted)
                {
                    wordDoc.MainDocumentPart.Document.Save();
                }
                else
                {
                    Console.WriteLine("No QR code images were inserted.");
                }
            }
        }

        public static byte[] GenerateQRCodeAsByteArray(string qrCodeText)
        {
            using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
            {
                // Create QR Code data
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(qrCodeText, QRCodeGenerator.ECCLevel.Q);
                using (QRCode qrCode = new QRCode(qrCodeData))
                using (Bitmap qrCodeImage = qrCode.GetGraphic(20))
                using (MemoryStream ms = new MemoryStream())
                {
                    qrCodeImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    return ms.ToArray();
                }
            }
        }


        ///////////////////////////////////////////Helper Codes///////////////////////////////////////////
        public static string SaveImageToTempFile(byte[] imageData)
        {
            try
            {
                string tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
                File.WriteAllBytes(tempFilePath, imageData);
                return tempFilePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving image to temporary file: {ex.Message}");
                return "error";
            }
        }

        private static bool InsertImageAtBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string imagePath)
        {
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return false;
            }

            // Determine if the bookmark is in a header, footer, or main document
            OpenXmlPart targetPart;
            if (bookmarkStart.Parent.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Header>().Any())
            {
                targetPart = wordDoc.MainDocumentPart.HeaderParts
                    .FirstOrDefault(h => h.RootElement.Descendants<BookmarkStart>().Contains(bookmarkStart));
            }
            else if (bookmarkStart.Parent.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Footer>().Any())
            {
                targetPart = wordDoc.MainDocumentPart.FooterParts
                    .FirstOrDefault(f => f.RootElement.Descendants<BookmarkStart>().Contains(bookmarkStart));
            }
            else
            {
                targetPart = wordDoc.MainDocumentPart; // Default to main document
            }

            if (targetPart == null)
            {
                Console.WriteLine("Unable to determine the target part for the bookmark.");
                return false;
            }

            // Add the image to the determined part
            ImagePart imagePart = targetPart switch
            {
                MainDocumentPart mainPart => mainPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png),
                HeaderPart headerPart => headerPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png),
                FooterPart footerPart => footerPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png),
                _ => throw new InvalidOperationException("Unsupported target part.")
            };

            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            // Define image dimensions
            const long width = 120 * 914400 / 96; // Adjust as needed
            const long height = 70 * 914400 / 96;

            // Create the drawing element
            var drawingElement = CreateImageElement(((OpenXmlPartContainer)targetPart).GetIdOfPart(imagePart), width, height);

            // Insert the image in the bookmark's context
            var runElement = bookmarkStart.Parent.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().FirstOrDefault();
            if (runElement == null)
            {
                runElement = new DocumentFormat.OpenXml.Wordprocessing.Run();
                bookmarkStart.InsertAfterSelf(runElement);
            }

            runElement.Append(drawingElement);

            Console.WriteLine($"Image successfully inserted at bookmark: {bookmarkStart.Name}");
            return true;
        }

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

        public static (BookmarkStart? bookmarkStart, BookmarkEnd? bookmarkEnd)? FindBookmark(WordprocessingDocument wordDoc, string bookmarkName)
        {
            // Retrieve all BookmarkStart elements (from the main document, headers, footers, etc.)
            var allBookmarks = wordDoc.MainDocumentPart.Document
                .Descendants<BookmarkStart>()
                .Concat(wordDoc.MainDocumentPart.HeaderParts
                    .SelectMany(header => header.RootElement.Descendants<BookmarkStart>()))
                .Concat(wordDoc.MainDocumentPart.FooterParts
                    .SelectMany(footer => footer.RootElement.Descendants<BookmarkStart>()))
                .Concat(wordDoc.MainDocumentPart.FootnotesPart?.RootElement
                    .Descendants<BookmarkStart>() ?? Enumerable.Empty<BookmarkStart>())
                .Concat(wordDoc.MainDocumentPart.EndnotesPart?.RootElement
                    .Descendants<BookmarkStart>() ?? Enumerable.Empty<BookmarkStart>());

            // Find the first BookmarkStart element that matches the specified bookmark name
            var bookmarkStart = allBookmarks.FirstOrDefault(b => b.Name == bookmarkName);
            if (bookmarkStart != null)
            {
                // Call the method to find the BookmarkEnd using the bookmarkStart.Id

                var bookmarkEnd = FindBookmarkEnd(wordDoc, bookmarkStart.Id);

                if (bookmarkEnd != null)
                {
                    // Return both BookmarkStart and BookmarkEnd as a tuple if both are found
                    return (bookmarkStart, bookmarkEnd);
                }
            }

            // Return null if no matching BookmarkStart or BookmarkEnd was found
            return null;
        }

        public static BookmarkEnd? FindBookmarkEnd(WordprocessingDocument wordDoc, string bookmarkId)
        {
            // Search each part for the first matching BookmarkEnd and return as soon as one is found
            return wordDoc.MainDocumentPart.Document.Descendants<BookmarkEnd>().FirstOrDefault(b => b.Id == bookmarkId)
                ?? wordDoc.MainDocumentPart.HeaderParts
                    .SelectMany(header => header.RootElement.Descendants<BookmarkEnd>())
                    .FirstOrDefault(b => b.Id == bookmarkId)
                ?? wordDoc.MainDocumentPart.FooterParts
                    .SelectMany(footer => footer.RootElement.Descendants<BookmarkEnd>())
                    .FirstOrDefault(b => b.Id == bookmarkId)
                ?? wordDoc.MainDocumentPart.FootnotesPart?.RootElement
                    .Descendants<BookmarkEnd>().FirstOrDefault(b => b.Id == bookmarkId)
                ?? wordDoc.MainDocumentPart.EndnotesPart?.RootElement
                    .Descendants<BookmarkEnd>().FirstOrDefault(b => b.Id == bookmarkId);
        }

        public static void EmptyBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, BookmarkEnd bookmarkEnd)
        {
            var currentElement = bookmarkStart.NextSibling();

            while (currentElement != null && currentElement != bookmarkEnd)
            {
                var nextElement = currentElement.NextSibling();

                // Remove text or image content while preserving styles
                if (currentElement is DocumentFormat.OpenXml.Wordprocessing.Run runElement)
                {
                    // Remove text elements
                    var textElements = runElement.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>().ToList();
                    foreach (var text in textElements)
                    {
                        text.Remove();
                    }

                    // Remove drawing/image elements
                    var drawingElements = runElement.Elements<DocumentFormat.OpenXml.Wordprocessing.Drawing>().ToList();
                    foreach (var drawing in drawingElements)
                    {
                        drawing.Remove();
                    }


                    // If the run becomes empty, retain it for style preservation
                    if (!runElement.HasChildren)
                    {
                        runElement.Remove();
                    }
                }

                currentElement = nextElement;
            }
        }




        //public static void ReplaceImageInDocument(string docxFilePath, string oldImageName, string newImagePath)
        //{
        //    // Open the Word document
        //    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docxFilePath, true))
        //    {
        //        // Get the main document part
        //        var mainPart = wordDoc.MainDocumentPart;

        //        // Check if the old image exists in the media folder
        //        var mediaFolder = Path.Combine(Path.GetDirectoryName(docxFilePath), "word", "media");

        //        string oldImagePath = Path.Combine(mediaFolder, oldImageName);
        //        if (!File.Exists(oldImagePath))
        //        {
        //            Console.WriteLine($"Old image '{oldImageName}' not found in the media folder.");
        //            return;
        //        }

        //        // Replace the old image with the new one
        //        string newImageName = Path.GetFileName(newImagePath);
        //        string newImagePathInMedia = Path.Combine(mediaFolder, newImageName);

        //        // Overwrite the old image with the new one
        //        File.Copy(newImagePath, newImagePathInMedia, true);

        //        // Update the image relationships in the document's XML
        //        UpdateImageRelationships(wordDoc, oldImageName, newImageName);

        //        // Update any bookmarks referencing this image
        //        UpdateBookmarkImageReferences(wordDoc, oldImageName, newImageName);

        //        Console.WriteLine("Image replacement and relationship updates completed.");
        //    }
        //}

        //private static void UpdateImageRelationships(WordprocessingDocument wordDoc, string oldImageName, string newImageName)
        //{
        //    // Path to the relationships file for the document
        //    var relsFilePath = wordDoc.ExtendedPropertiesPart.Uri.OriginalString;
        //    XDocument relsDoc = XDocument.Load(relsFilePath);

        //    // Find and update the old image's relationship
        //    var imageRelationship = relsDoc.Descendants().FirstOrDefault(
        //        r => r.Name.LocalName == "Relationship" && r.Attribute("Target").Value.Contains(oldImageName));

        //    if (imageRelationship != null)
        //    {
        //        imageRelationship.SetAttributeValue("Target", "media/" + newImageName);
        //        relsDoc.Save(relsFilePath);
        //    }
        //}

        //private static void UpdateBookmarkImageReferences(WordprocessingDocument wordDoc, string oldImageName, string newImageName)
        //{
        //    // Traverse through all bookmarks and their images
        //    var bookmarks = wordDoc.MainDocumentPart.Document.Descendants<BookmarkStart>();

        //    foreach (var bookmark in bookmarks)
        //    {
        //        // Find runs with embedded images at the bookmark location
        //        var runsWithImages = bookmark.Parent.Descendants<Run>().Where(r =>
        //            r.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any(d =>
        //                d.Descendants < DocumentFormat.OpenXml.Drawing.Pictures.Blip).Any(b =>
        //                    b.Embed.Value.Contains(oldImageName))).ToList());

        //        foreach (var run in runsWithImages)
        //        {
        //            var drawing = run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().FirstOrDefault();
        //            var blip = drawing?.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Blip>().FirstOrDefault();

        //            if (blip != null)
        //            {
        //                // Update the reference to the new image
        //                blip.Embed = newDocumentPart.GetIdOfPart(newImageName);
        //            }
        //        }
        //    }
        //}


    }
}
