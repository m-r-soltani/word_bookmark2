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
                // Find the parent element of the bookmark
                var parentElement = bookmarkStart.Parent;

                if (parentElement == null)
                {
                    Console.WriteLine("Parent element is null. Unable to insert text.");
                    return false;
                }

                // Create a new Run element for the bookmark's content
                var run = new DocumentFormat.OpenXml.Wordprocessing.Run();

                // Split the new text by line breaks
                var lines = newText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                // Append Text and Break elements for each line
                for (int i = 0; i < lines.Length; i++)
                {
                    // Add a Text element for the current line
                    if (!string.IsNullOrEmpty(lines[i]))
                    {
                        var textElement = new DocumentFormat.OpenXml.Wordprocessing.Text(lines[i])
                        {
                            Space = SpaceProcessingModeValues.Preserve // Preserve spaces
                        };
                        run.AppendChild(textElement);
                    }

                    // Add a Break element if it's not the last line
                    if (i < lines.Length - 1)
                    {
                        run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Break());
                    }
                }

                // Insert the Run after the BookmarkStart
                bookmarkStart.InsertAfterSelf(run);

                // Text insertion successful
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during text insertion: {ex.Message}");
                return false;
            }
        }


        public static bool InsertTextAtBookmark_likelastonelinebreaknotgood5(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string newText)
        {
            try
            {
                // Validate the parent element
                var parentElement = bookmarkStart.Parent;
                if (parentElement == null)
                {
                    Console.WriteLine("Parent element is null. Unable to insert text.");
                    return false;
                }

                // Split the new text by line breaks
                var lines = newText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                // Iterate over the lines and create paragraphs for each
                foreach (var line in lines)
                {
                    // Create a new paragraph for each line
                    var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();

                    // Add a Run and Text element with the line's content
                    if (!string.IsNullOrEmpty(line))
                    {
                        var run = new DocumentFormat.OpenXml.Wordprocessing.Run();
                        var textElement = new DocumentFormat.OpenXml.Wordprocessing.Text(line)
                        {
                            Space = SpaceProcessingModeValues.Preserve // Preserve spaces
                        };
                        run.AppendChild(textElement);
                        paragraph.AppendChild(run);
                    }

                    // Insert the paragraph after the BookmarkStart's parent
                    parentElement.InsertAfterSelf(paragraph);
                }

                // Text insertion successful
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during text insertion: {ex.Message}");
                return false;
            }
        }


        public static bool InsertTextAtBookmark_simplified_linebreak_stretches4(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string newText)
        {
            try
            {
                // Validate the parent element
                var parentElement = bookmarkStart.Parent;
                if (parentElement == null)
                {
                    Console.WriteLine("Parent element is null. Unable to insert text.");
                    return false;
                }

                // Create a new Run element
                var run = new DocumentFormat.OpenXml.Wordprocessing.Run();

                // Split the new text by line breaks and add Text and Break elements
                var lines = newText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                for (int i = 0; i < lines.Length; i++)
                {
                    // Add the text element
                    if (!string.IsNullOrEmpty(lines[i]))
                    {
                        var textElement = new DocumentFormat.OpenXml.Wordprocessing.Text(lines[i])
                        {
                            Space = SpaceProcessingModeValues.Preserve // Preserve spaces
                        };
                        run.AppendChild(textElement);
                    }

                    // Add a break if it's not the last line
                    if (i < lines.Length - 1)
                    {
                        run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Break());
                    }
                }

                // Insert the Run element after the BookmarkStart
                parentElement.InsertAfter(run, bookmarkStart);

                // Text insertion successful
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during text insertion: {ex.Message}");
                return false;
            }
        }


        public static bool InsertTextAtBookmark_workswell3(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string newText)
        {
            try
            {
                // Find the parent element of the bookmark
                var parentElement = bookmarkStart.Parent;

                if (parentElement is DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
                {
                    // Create a new Run element
                    var run = new DocumentFormat.OpenXml.Wordprocessing.Run();

                    // Split the new text by line breaks and add Text and Break elements
                    var lines = newText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (!string.IsNullOrEmpty(lines[i]))
                        {
                            var textElement = new DocumentFormat.OpenXml.Wordprocessing.Text(lines[i])
                            {
                                Space = SpaceProcessingModeValues.Preserve // Preserve spaces
                            };
                            run.AppendChild(textElement);
                        }

                        if (i < lines.Length - 1)
                        {
                            run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Break()); // Add a break for all but the last line
                        }
                    }

                    // Insert the Run after the BookmarkStart
                    paragraph.InsertAfter(run, bookmarkStart);
                }
                else if (parentElement != null)
                {
                    // Handle non-paragraph parents (e.g., tables)
                    var run = new DocumentFormat.OpenXml.Wordprocessing.Run();
                    var lines = newText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                    foreach (var line in lines)
                    {
                        if (!string.IsNullOrEmpty(line))
                        {
                            var textElement = new DocumentFormat.OpenXml.Wordprocessing.Text(line)
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            };
                            run.AppendChild(textElement);
                        }
                        run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Break());
                    }

                    parentElement.InsertAfter(run, bookmarkStart);
                }
                else
                {
                    Console.WriteLine("Parent element is null. Unable to insert text.");
                    return false;
                }

                // Text insertion successful
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during text insertion: {ex.Message}");
                return false;
            }
        }


        public static bool InsertTextAtBookmark2(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string newText)
        {
            try
            {
                // Find the parent element of the bookmark
                var parentElement = bookmarkStart.Parent;

                if (parentElement is DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
                {
                    // Find the first Run after the bookmark to clone its style (if needed)
                    var existingRun = paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().FirstOrDefault();
                    var run = existingRun != null
                        ? (DocumentFormat.OpenXml.Wordprocessing.Run)existingRun.CloneNode(true)
                        : new DocumentFormat.OpenXml.Wordprocessing.Run();

                    // Remove any existing text from the cloned Run
                    run.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Text>();
                    run.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Break>();

                    // Split the new text by line breaks and create Text and Break elements
                    var lines = newText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        // Append a Text element for each line
                        run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(lines[i]));

                        // Add a Break element after each line except the last one
                        if (i < lines.Length - 1)
                        {
                            run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Break());
                        }
                    }

                    // Insert the Run after the BookmarkStart
                    paragraph.InsertAfter(run, bookmarkStart);
                }
                else if (parentElement != null)
                {
                    // Handle non-paragraph parents (e.g., tables)
                    var run = new DocumentFormat.OpenXml.Wordprocessing.Run();

                    // Split the new text by line breaks and add Text and Break elements
                    var lines = newText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(lines[i]));
                        if (i < lines.Length - 1)
                        {
                            run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Break());
                        }
                    }

                    parentElement.InsertAfter(run, bookmarkStart);
                }
                else
                {
                    // Parent element is null
                    Console.WriteLine("Parent element is null. Unable to insert text.");
                    return false;
                }

                // Text insertion successful
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during text insertion: {ex.Message}");
                return false;
            }
        }


        public static bool InsertTextAtBookmark1(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string newText)
        {

            try
            {
                // Find the parent element of the bookmark
                var parentElement = bookmarkStart.Parent;

                if (parentElement is DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
                {
                    // Find the first Run after the bookmark to clone its style (if needed)
                    var existingRun = paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().FirstOrDefault();
                    if (existingRun != null)
                    {
                        Console.WriteLine("Found an existing Run to clone style.");
                    }
                    else
                    {
                        Console.WriteLine("No existing Run found. Creating a new one.");
                    }
                    var run = existingRun != null
                        ? (DocumentFormat.OpenXml.Wordprocessing.Run)existingRun.CloneNode(true)
                        : new DocumentFormat.OpenXml.Wordprocessing.Run();

                    // Debug: Log that we're removing existing text from the cloned Run
                    Console.WriteLine("Removing existing text from the Run.");
                    run.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Text>();

                    // Debug: Log that we're appending the new text
                    Console.WriteLine("Appending new text to the Run.");
                    run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(newText));

                    // Debug: Log that we're inserting the Run after the bookmark
                    Console.WriteLine("Inserting the Run after the BookmarkStart.");
                    paragraph.InsertAfter(run, bookmarkStart);
                }
                else if (parentElement != null)
                {
                    // Handle non-paragraph parents (e.g., tables)
                    var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(newText));

                    // Debug: Log that we're inserting the Run after the bookmark
                    Console.WriteLine("Inserting the Run after the BookmarkStart.");
                    parentElement.InsertAfter(run, bookmarkStart);
                }
                else
                {
                    // Debug: Log that the parent element is null
                    Console.WriteLine("Parent element is null. Unable to insert text.");
                    return false; // Parent element couldn't be determined
                }

                // Debug: Log success
                Console.WriteLine("Text insertion successful.");
                return true; // Insertion successful
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during text insertion, Message: "+ex.Message + " StackTrace: " + ex.StackTrace);
                return false; // Insertion failed due to an exception
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

            // Determine the parent element and part where the image will be inserted
            var mainPart = wordDoc.MainDocumentPart;
            OpenXmlPart targetPart = null;
            var parentElement = bookmarkStart.Parent;

            // Determine if the bookmark is inside a header, footer, or main document
            if (parentElement.Ancestors<Header>().Any())
            {
                targetPart = wordDoc.MainDocumentPart.HeaderParts
                    .FirstOrDefault(h => h.RootElement.Descendants<BookmarkStart>().Contains(bookmarkStart));
            }
            else if (parentElement.Ancestors<Footer>().Any())
            {
                targetPart = wordDoc.MainDocumentPart.FooterParts
                    .FirstOrDefault(f => f.RootElement.Descendants<BookmarkStart>().Contains(bookmarkStart));
            }
            else
            {
                targetPart = mainPart;
            }

            if (targetPart == null)
            {
                Console.WriteLine("Unable to determine the target part for the bookmark.");
                return false;
            }

            // Add the image part to the determined part
            ImagePart imagePart = null;
            if (targetPart is MainDocumentPart mainDocPart)
            {
                imagePart = mainDocPart.AddImagePart(ImagePartType.Png);
            }
            else if (targetPart is HeaderPart headerPart)
            {
                imagePart = headerPart.AddImagePart(ImagePartType.Png);
            }
            else if (targetPart is FooterPart footerPart)
            {
                imagePart = footerPart.AddImagePart(ImagePartType.Png);
            }

            if (imagePart == null)
            {
                Console.WriteLine("Failed to add an image part.");
                return false;
            }

            // Load the image data into the image part
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            // Get image dimensions (you can calculate this based on the actual image dimensions if needed)
            const long width = 120 * 914400 / 96; // Adjust width in EMUs
            const long height = 70 * 914400 / 96; // Adjust height in EMUs

            // Create the image element
            var drawingElement = CreateImageElement(((OpenXmlPartContainer)targetPart).GetIdOfPart(imagePart), width, height);

            // Insert the Run containing the image at the bookmark
            var runElement = new DocumentFormat.OpenXml.Wordprocessing.Run(drawingElement);
            parentElement.InsertAfter(runElement, bookmarkStart);

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
                var nextElement = currentElement.NextSibling(); // Cache next sibling

                // Skip removing nested bookmarks' start and end
                if (!(currentElement is BookmarkStart || currentElement is BookmarkEnd))
                {
                    currentElement.Remove(); // Remove only non-bookmark elements
                }

                currentElement = nextElement; // Move to the next sibling
            }
        }


        //public static void EmptyBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, BookmarkEnd bookmarkEnd)
        //{
        //    // Start from the sibling after the BookmarkStart element and iterate until the BookmarkEnd
        //    var currentElement = bookmarkStart.NextSibling();
        //    while (currentElement != null && currentElement != bookmarkEnd)
        //    {
        //        var nextElement = currentElement.NextSibling(); // Cache next sibling
        //        currentElement.Remove(); // Remove current element
        //        currentElement = nextElement; // Move to the next sibling
        //    }
        //}


        //public static void EmptyBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, BookmarkEnd bookmarkEnd)
        //{
        //    // List to hold elements to remove
        //    var elementsToRemove = new List<OpenXmlElement>();

        //    // Start from the sibling after the BookmarkStart element and iterate until the BookmarkEnd
        //    for (var currentElement = bookmarkStart.NextSibling(); currentElement != null && currentElement != bookmarkEnd; currentElement = currentElement.NextSibling())
        //    {
        //        // Collect all elements in the range to be removed
        //        elementsToRemove.Add(currentElement);

        //        // Recursively collect child elements if needed (for nested elements)
        //        CollectNestedElements(currentElement, elementsToRemove);
        //    }

        //    // Remove all collected elements
        //    foreach (var element in elementsToRemove)
        //    {
        //        // Safety check in case something is unexpectedly null
        //        if (element != null)
        //        {
        //            element.Remove();
        //        }
        //    }
        //}

        //// Helper method to remove child elements recursively in EmptyBookmark method
        //private static void CollectNestedElements(OpenXmlElement element, List<OpenXmlElement> elementsToRemove)
        //{
        //    if (element == null) return; // Safety check for null

        //    // Loop through all child elements and add them to the list to be removed
        //    foreach (var child in element.Elements())
        //    {
        //        if (child == null) continue; // Safety check for null child

        //        elementsToRemove.Add(child);
        //        // Recursively collect child elements if they have their own children
        //        CollectNestedElements(child, elementsToRemove);
        //    }
        //}


















































        //refactored simple inserttextatbookmark
        //public static void InsertTextAtBookmark2(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string newText)
        //{
        //    // Find the parent element of the bookmark
        //    var parentElement = bookmarkStart.Parent;

        //    if (parentElement is DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
        //    {
        //        // Insert a new Run containing the text directly into the existing Paragraph
        //        var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(newText));
        //        paragraph.InsertAfter(run, bookmarkStart);
        //    }
        //    else if (parentElement != null)
        //    {
        //        // For non-paragraph parents (like tables), add the new content in the parent
        //        var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(newText));
        //        parentElement.InsertAfter(run, bookmarkStart);
        //    }
        //    else
        //    {
        //        throw new InvalidOperationException("The bookmark's parent element could not be determined.");
        //    }
        //}

        //function made inserttextbookmark
        //public static void InsertTextAtBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string newText)
        //{
        //    // Create a new run with the new text content (use Wordprocessing namespace)
        //    var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(newText));

        //    // Create a new paragraph to hold the new run (use Wordprocessing namespace)
        //    var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(run);

        //    // Find the parent element (usually a paragraph) of the BookmarkStart
        //    var parentElement = bookmarkStart.Parent as OpenXmlElement;

        //    if (parentElement != null)
        //    {
        //        // Insert the new content at the position where the bookmark is located
        //        parentElement.InsertAfter(paragraph, bookmarkStart);
        //    }

        //}


    }
}
