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

namespace BookMarks
{
    public class BookmarkOpenxml
    {
        public static void UpdateBookmarks(string filePath, Dictionary<string, string> bookmarksContent)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                foreach (var entry in bookmarksContent)
                {
                    string bookmarkName = entry.Key;
                    string content = entry.Value;

                    // Find the bookmark (this will give you the BookmarkStart)
                    var bookmarkStart = FindBookmark(wordDoc, bookmarkName);
                    //Console.WriteLine(GetType(BookmarkStart));
                    if (bookmarkStart != null)
                    {
                        // Find the corresponding BookmarkEnd using the same Id
                        var bookmarkEnd = wordDoc.MainDocumentPart.Document.Body
                            .Descendants<BookmarkEnd>()
                            .FirstOrDefault(be => be.Id == bookmarkStart.Id);

                        if (bookmarkEnd != null)
                        {
                            // Remove content between BookmarkStart and BookmarkEnd
                            var elementsToRemove = new List<OpenXmlElement>();
                            for (var currentElement = bookmarkStart.NextSibling(); currentElement != null && currentElement != bookmarkEnd; currentElement = currentElement.NextSibling())
                            {
                                elementsToRemove.Add(currentElement);
                            }

                            // Remove the elements found between BookmarkStart and BookmarkEnd
                            foreach (var element in elementsToRemove)
                            {
                                element.Remove();
                            }

                            // Insert new content (text or image) at the bookmark
                            if (content.StartsWith("@QRCode:"))
                            {
                                string qrCodeText = content.Substring(8);
                                string qrCodeImagePath = GenerateQRCodeImage(qrCodeText);
                                InsertImageAtBookmark(wordDoc, bookmarkStart, qrCodeImagePath);
                            }
                            else if (File.Exists(content))
                            {
                                InsertImageAtBookmark(wordDoc, bookmarkStart, content);
                            }
                            else if (content.StartsWith("@Binary:"))
                            {
                                string binaryFilePath = ProcessBinaryImage(new Dictionary<string, string> { { entry.Key, entry.Value } });
                                if (binaryFilePath != "error")
                                {
                                    InsertImageAtBookmark(wordDoc, bookmarkStart, binaryFilePath);
                                }
                            }
                            else
                            {
                                // Insert new text at the bookmark
                                var newRun = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(content));
                                bookmarkStart.Parent.InsertAfter(newRun, bookmarkStart);
                            }
                        }
                        else
                        {
                            Console.WriteLine($"BookmarkEnd for '{bookmarkName}' not found.");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Bookmark '{bookmarkName}' not found in the document.");
                    }
                }

                // Save the document after updating all bookmarks
                wordDoc.MainDocumentPart.Document.Save();
            }
        }
  
        /*before any changes
        public static void UpdateBookmarks(string filePath, Dictionary<string, string> bookmarksContent)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                foreach (var entry in bookmarksContent)
                {
                    string bookmarkName = entry.Key;
                    string content = entry.Value;

                    // Find the bookmark
                    var bookmark = FindBookmark(wordDoc, bookmarkName);

                    if (bookmark != null)
                    {
                        // Locate the bookmark's text content
                        var nextSibling = bookmark.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Run>();

                        if (content.StartsWith("@QRCode:"))
                        {

                            // QR code content starts with @QRCode: (e.g., @QRCode:some_text)
                            string qrCodeText = content.Substring(8); // Extract the QR code text
                            string qrCodeImagePath = GenerateQRCodeImage(qrCodeText); // Generate the QR code image
                            InsertImageAtBookmark(wordDoc, bookmark, qrCodeImagePath);
                        }
                        else if (File.Exists(content)) // Check if the value is an image file path
                        {
                            // Remove existing content (if any)
                            nextSibling?.Remove();
                            InsertImageAtBookmark(wordDoc, bookmark, content);
                        } else if (content.StartsWith("@Binary:")) { // Check if the value is an image binary
                            string thisfilepath=ProcessBinaryImage(new Dictionary<string, string> { { entry.Key, entry.Value } });
                            if (thisfilepath != "error") {
                                InsertImageAtBookmark(wordDoc, bookmark, thisfilepath);
                            }
                        }
                        else // Treat the value as text
                        {

                            if (nextSibling != null)
                            {
                                var textElement = nextSibling.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>();
                                if (textElement != null)
                                {
                                    textElement.Text = content; // Update the text
                                }
                            }
                            else
                            {
                                // If no sibling, insert a new Run with the text
                                var newRun = new DocumentFormat.OpenXml.Wordprocessing.Run(
                                    new DocumentFormat.OpenXml.Wordprocessing.Text(content));
                                bookmark.Parent.InsertAfter(newRun, bookmark);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Bookmark '{bookmarkName}' not found in the document.");
                    }
                }

                // Save the document after updating all bookmarks
                wordDoc.MainDocumentPart.Document.Save();
            }
        }
        */

        private static void EmptyBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, BookmarkEnd bookmarkEnd) {
            // Manually iterate siblings to collect elements between BookmarkStart and BookmarkEnd
            var elementsToRemove = new List<OpenXmlElement>();
            for (var currentElement = bookmarkStart.NextSibling(); currentElement != null && currentElement != bookmarkEnd; currentElement = currentElement.NextSibling())
            {
                elementsToRemove.Add(currentElement);
            }

            // Remove the collected elements
            foreach (var element in elementsToRemove)
            {
                element.Remove();
            }
        }

        public static BookmarkStart FindBookmark(WordprocessingDocument wordDoc, string bookmarkName)
        {
            // Search bookmarks in the main document body
            var bookmarks = wordDoc.MainDocumentPart.Document.Descendants<BookmarkStart>()
                .Where(b => b.Name == bookmarkName);

            if (bookmarks.Any())
                return bookmarks.First();

            // Search in headers
            foreach (var header in wordDoc.MainDocumentPart.HeaderParts)
            {
                var headerBookmarks = header.RootElement.Descendants<BookmarkStart>()
                    .Where(b => b.Name == bookmarkName);
                if (headerBookmarks.Any())
                    return headerBookmarks.First();
            }
            // Search in footers
            foreach (var footer in wordDoc.MainDocumentPart.FooterParts)
            {
                var footerBookmarks = footer.RootElement.Descendants<BookmarkStart>()
                    .Where(b => b.Name == bookmarkName);
                if (footerBookmarks.Any())
                    return footerBookmarks.First();
            }

            // Search in footnoteBookmarks
            if (wordDoc.MainDocumentPart.FootnotesPart != null)
            {
                var footnoteBookmarks = wordDoc.MainDocumentPart.FootnotesPart.RootElement.Descendants<BookmarkStart>()
                    .Where(b => b.Name == bookmarkName);
                if (footnoteBookmarks.Any())
                    return footnoteBookmarks.First();
            }
            // Search in endnoteBookmarks
            if (wordDoc.MainDocumentPart.EndnotesPart != null)
            {
                var endnoteBookmarks = wordDoc.MainDocumentPart.EndnotesPart.RootElement.Descendants<BookmarkStart>()
                    .Where(b => b.Name == bookmarkName);
                if (endnoteBookmarks.Any())
                    return endnoteBookmarks.First();
            }

            // If not found, return null
            return null;
        }

        private static void InsertImageAtBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmark, string imagePath)
        {
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            var mainPart = wordDoc.MainDocumentPart;
            var parentElement = bookmark.Parent;

            // Add the image to the correct part (MainDocumentPart or Header/Footer)
            ImagePart imagePart = null;

            if (parentElement.Ancestors<Header>().Any())
            {
                var headerPart = wordDoc.MainDocumentPart.HeaderParts.FirstOrDefault(h => h.RootElement.Descendants<BookmarkStart>().Contains(bookmark));
                if (headerPart != null)
                {
                    imagePart = headerPart.AddImagePart(ImagePartType.Png);
                }
            }
            else if (parentElement.Ancestors<Footer>().Any())
            {
                var footerPart = wordDoc.MainDocumentPart.FooterParts.FirstOrDefault(f => f.RootElement.Descendants<BookmarkStart>().Contains(bookmark));
                if (footerPart != null)
                {
                    imagePart = footerPart.AddImagePart(ImagePartType.Png);
                }
            }
            else
            {
                // Default to the main document part
                imagePart = mainPart.AddImagePart(ImagePartType.Png);
            }

            // Load the image data
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            // Get image dimensions
            const long width = 120 * 914400 / 96;
            const long height = 70 * 914400 / 96;

            // Create the image element
            var drawingElement = CreateImageElement(mainPart.GetIdOfPart(imagePart), width, height);

            var imageRun = new DocumentFormat.OpenXml.Wordprocessing.Run(drawingElement);

            // Insert the image at the bookmark
            parentElement.InsertAfter(imageRun, bookmark);
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

        public static string ProcessBinaryImage(Dictionary<string, string> bookmarksContent)
        {
            var key = bookmarksContent.Keys.First(); // Get the single key
            var value = bookmarksContent[key];
            string userName = value.Substring(8); // Extract user name
            byte[] imageData = GetImageDataFromDatabase(userName); // Fetch binary data
            if (imageData != null)
            {
                string tempFilePath = SaveImageToTempFile(imageData);
                if (File.Exists(tempFilePath))
                {
                    return tempFilePath;
                }
                else {
                    return "error";
                }
            }
            else {
                return "error";
            }
        }

        public static string SaveImageToTempFile(byte[] imageData)
        {
            string tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
            Console.WriteLine($"Writing image data to: {tempFilePath}"); // Debug info
            File.WriteAllBytes(tempFilePath, imageData);
            return tempFilePath;
        }

        public static string GenerateQRCodeImage(string qrCodeText)
        {
            // Create a new instance of the QRCodeGenerator
            using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
            {
                // Create a QR Code data
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(qrCodeText, QRCodeGenerator.ECCLevel.Q);

                // Create a QRCode object from the data
                using (QRCode qrCode = new QRCode(qrCodeData))
                {
                    // Create the QR code image
                    using (Bitmap qrCodeImage = qrCode.GetGraphic(20)) // The '20' is the pixel size multiplier for the code
                    {
                        // Save the image to a temporary file
                        string tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
                        qrCodeImage.Save(tempFilePath, System.Drawing.Imaging.ImageFormat.Png);

                        return tempFilePath; // Return the file path to the generated QR code
                    }
                }
            }
        }







        //private static (long width, long height) GetImageDimensions(string imagePath, long maxWidthEmu = 3000000, long maxHeightEmu = 3000000)
        //{
        //    using (var image = SixLabors.ImageSharp.Image.Load(imagePath))
        //    {
        //        const int emusPerInch = 914400; // Conversion factor

        //        var dpiX = image.Metadata.HorizontalResolution > 0 ? image.Metadata.HorizontalResolution : 96; // Default DPI
        //        var dpiY = image.Metadata.VerticalResolution > 0 ? image.Metadata.VerticalResolution : 96;

        //        var widthEmu = (long)(image.Width * emusPerInch / dpiX);
        //        var heightEmu = (long)(image.Height * emusPerInch / dpiY);

        //        // Scale to fit within max dimensions
        //        if (widthEmu > maxWidthEmu || heightEmu > maxHeightEmu)
        //        {
        //            double widthScale = (double)maxWidthEmu / widthEmu;
        //            double heightScale = (double)maxHeightEmu / heightEmu;
        //            double scale = Math.Min(widthScale, heightScale);

        //            widthEmu = (long)(widthEmu * scale);
        //            heightEmu = (long)(heightEmu * scale);
        //        }

        //        return (widthEmu, heightEmu);
        //    }
        //}

        //without qrcode
        /*
        public static void UpdateBookmarks(string filePath, Dictionary<string, string> bookmarksContent)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                foreach (var entry in bookmarksContent)
                {
                    string bookmarkName = entry.Key;
                    string content = entry.Value;

                    // Find the bookmark
                    var bookmark = FindBookmark(wordDoc, bookmarkName);

                    if (bookmark != null)
                    {
                        // Locate the bookmark's text content
                        var nextSibling = bookmark.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Run>();

                        if (File.Exists(content)) // Check if the value is an image file path
                        {
                            // Remove existing content (if any)
                            nextSibling?.Remove();

                            // Add the image
                            InsertImageAtBookmark(wordDoc, bookmark, content);
                        }
                        else // Treat the value as text
                        {
                            if (nextSibling != null)
                            {
                                var textElement = nextSibling.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>();
                                if (textElement != null)
                                {
                                    textElement.Text = content; // Update the text
                                }
                            }
                            else
                            {
                                // If no sibling, insert a new Run with the text
                                var newRun = new DocumentFormat.OpenXml.Wordprocessing.Run(
                                    new DocumentFormat.OpenXml.Wordprocessing.Text(content));
                                bookmark.Parent.InsertAfter(newRun, bookmark);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Bookmark '{bookmarkName}' not found in the document.");
                    }
                }

                // Save the document after updating all bookmarks
                wordDoc.MainDocumentPart.Document.Save();
            }
        }
        */

        /*
        private static (long width, long height) GetImageDimensions(string imagePath)
        {
            //using (var image = SixLabors.ImageSharp.Image.Load(imagePath)) // This is a more general method without the pixel format
            using (var image = SixLabors.ImageSharp.Image.Load(imagePath)) // This is a more general method without the pixel format            
            {
                const int emusPerInch = 914400; // Conversion factor

                var dpiX = image.Metadata.HorizontalResolution;
                var dpiY = image.Metadata.VerticalResolution;

                // Return dimensions in EMUs
                return (
                    (long)(image.Width * emusPerInch / dpiX),
                    (long)(image.Height * emusPerInch / dpiY)
                );
            }
        }

         */
        /*
        public static byte[] GetImageDataFromDatabase(string userName)
        {
            string connectionString = "Server=.;Database=CentralUserInfo;User Id=sa;Password=Aa@12345;";
            string query = "SELECT FIRST_SIGNATURE FROM CentralUserInfo.dbo.users WHERE user_name = @UserName";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
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
                    else
                    {
                        Console.WriteLine($"Unexpected data type for user: {userName}. Data type: {result.GetType()}");
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving data for user: {userName}. Exception: {ex.Message}");
                return null;
            }
        }
        */

        //GetImageDataFromDatabase checking base64

        /*
        public static byte[] GetImageDataFromDatabase(string userName)
        {
            string connectionString = "Server=.;Database=CentralUserInfo;User Id=sa;Password=Aa@12345;";
            string query = "SELECT FIRST_SIGNATURE FROM CentralUserInfo.dbo.users WHERE user_name = @UserName";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
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

                    if (result is string base64String) // Check if the data is a string
                    {
                        Console.WriteLine($"Base64 string data retrieved for user: {userName}");
                        return Convert.FromBase64String(base64String); // Decode Base64 string to binary
                    }

                    if (result is byte[] binaryData) // Check if the data is binary
                    {
                        Console.WriteLine($"Binary data retrieved for user: {userName}");
                        return binaryData;
                    }

                    Console.WriteLine($"Unexpected data type for user: {userName}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving data for user: {userName}. Exception: {ex.Message}");
                return null;
            }
        }
        */

        //GetImageDataFromDatabase varbinary
        /*
        public static byte[] GetImageDataFromDatabase(string userName)
        {
            string connectionString = "Server=.;Database=CentralUserInfo;User Id=sa;Password=Aa@12345;";
            string query = "SELECT FIRST_SIGNATURE FROM CentralUserInfo.dbo.users2 WHERE user_name = @UserName";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
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

                    Console.WriteLine($"Data retrieved for user: {userName}"); // Debug info
                    return (byte[])result;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving data for user: {userName}. Exception: {ex.Message}");
                return null;
            }
        }
        */





        /*
        public static byte[] GetImageDataFromDatabase(string userName)
        {
            //string connectionString = "YourDatabaseConnectionStringHere"; // Replace with your connection string
            string connectionString = "Server=.;Database=CentralUserInfo;User Id=sa;Password=Aa@12345;";
            string query = "SELECT FIRST_SIGNATURE FROM CentralUserInfo.dbo.users WHERE user_name = @UserName";

            using (SqlConnection connection = new SqlConnection(connectionString))
            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@UserName", userName);
                connection.Open();
                var result = command.ExecuteScalar();
                return result as byte[];
            }
        }
        */

        /*
        public static string SaveImageToTempFile(byte[] imageData)
        {
            string tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
            File.WriteAllBytes(tempFilePath, imageData);
            return tempFilePath;
        }
        */

        /*
        public static void ProcessBinaryImages(Dictionary<string, string> bookmarksContent)
        {
            foreach (var key in bookmarksContent.Keys)
            {
                if (bookmarksContent[key].StartsWith("@Binary:"))
                {
                    string userName = bookmarksContent[key].Substring(8); // Extract user name
                    byte[] imageData = GetImageDataFromDatabase(userName); // Fetch binary data
                    if (imageData != null)
                    {
                        string tempFilePath = SaveImageToTempFile(imageData);
                        bookmarksContent[key] = tempFilePath; // Replace marker with temp file path
                    }
                    else
                    {
                        Console.WriteLine($"No image data found for user: {userName}");
                        bookmarksContent[key] = string.Empty; // Clear the marker if no data found
                    }
                }
            }
        }
        */
        /*
        private static void InsertImageAtBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmark, string imagePath)
        {
            var mainPart = wordDoc.MainDocumentPart;
            var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            // Add the image to the document
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            // Create the image element
            //var drawingElement = CreateImageElement(mainPart.GetIdOfPart(imagePart), 3000000, 2000000); // Example size in EMUs
            var (width, height) = GetImageDimensions(imagePath);
            var drawingElement = CreateImageElement(mainPart.GetIdOfPart(imagePart), width, height);
            // Create a Run with the drawing
            var imageRun = new DocumentFormat.OpenXml.Wordprocessing.Run(drawingElement);

            // Insert the image after the bookmark
            bookmark.Parent.InsertAfter(imageRun, bookmark);
        }
        */














        /*
        private static void InsertImageAtBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmark, string imagePath)
        {
            var mainPart = wordDoc.MainDocumentPart;

            // Get the file extension in lowercase
            string fileExtension = System.IO.Path.GetExtension(imagePath).ToLowerInvariant();

            
            // Determine the ImagePartType based on the file extension
            //ImagePartType imagePartType;
            switch (fileExtension)
            {
                case ".jpg":
                case ".jpeg":
                    var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    break;

                case ".png":
                    var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                    break;

                case ".bmp":
                    imagePartType = ImagePartType.Bmp;
                    break;

                case ".gif":
                    imagePartType = ImagePartType.Gif;
                    break;

                case ".tiff":
                    imagePartType = ImagePartType.Tiff;
                    break;

                default:
                    throw new NotSupportedException($"Unsupported image type: {fileExtension}");
            }

            // Add the image part to the document
            //var imagePart = mainPart.AddImagePart(imagePartType);

            // Feed image data from the file to the image part
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            // Calculate image dimensions and create the drawing element
            var (width, height) = GetImageDimensions(imagePath);
            var drawingElement = CreateImageElement(mainPart.GetIdOfPart(imagePart), width, height);

            // Create a new run for the image and insert it after the bookmark
            var imageRun = new DocumentFormat.OpenXml.Wordprocessing.Run(drawingElement);
            bookmark.Parent.InsertAfter(imageRun, bookmark);
        }
        */

    }
}
