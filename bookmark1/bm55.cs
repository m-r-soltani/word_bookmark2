using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using QRCoder;
using System.Data.SqlClient;
using System.Drawing.Imaging;
using System.Drawing;
using System.IO.Compression;


namespace bm5
{
  
    
    public class bm5
    {
        // Main method to update bookmarks
        public static void UpdateBookmarks(
            string docPath,
            Dictionary<string, string> textBookmarkContents,
            Dictionary<string, byte[]> imageBookmarkContents,
            Dictionary<string, byte[]> qrcodeBookmarkContents)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docPath, true))
            {
                // Handle text bookmarks
                if (textBookmarkContents != null)
                {
                    foreach (var bookmark in textBookmarkContents)
                    {
                        UpdateTextBookmark(wordDoc, bookmark.Key, bookmark.Value);
                    }
                }

                // Handle image bookmarks
                if (imageBookmarkContents != null)
                {
                    foreach (var bookmark in imageBookmarkContents)
                    {
                        InsertImageBookmark(wordDoc, bookmark.Key, bookmark.Value);
                    }
                }

                // Handle QR code bookmarks
                if (qrcodeBookmarkContents != null)
                {
                    foreach (var bookmark in qrcodeBookmarkContents)
                    {
                        InsertQRCodeBookmark(wordDoc, bookmark.Key, bookmark.Value);
                    }
                }
            }
        }

        // Update text in a bookmark
        private static void UpdateTextBookmark(WordprocessingDocument wordDoc, string bookmarkName, string content)
        {
            var bookmark = FindBookmark(wordDoc, bookmarkName);
            if (bookmark != null)
            {
                var parent = bookmark.Parent;
                var textElement = parent.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().FirstOrDefault();
                if (textElement != null)
                {
                    textElement.Text = content;
                }
                else
                {
                    parent.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(content));
                }
            }
        }

        // Insert an image at a bookmark
        private static void InsertImageBookmark(WordprocessingDocument wordDoc, string bookmarkName, byte[] imageBytes)
        {
            var bookmark = FindBookmark(wordDoc, bookmarkName);
            if (bookmark != null && ValidateImageData(imageBytes))
            {
                var parent = bookmark.Parent;
                var imageRun = CreateImageRun(wordDoc, imageBytes, "image.png", 110, 60);

                // Check if the bookmark is in a header or footer
                if (bookmark.Ancestors<Header>().Any())
                {
                    var header = bookmark.Ancestors<Header>().First();
                    header.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(imageRun));
                }
                else if (bookmark.Ancestors<Footer>().Any())
                {
                    var footer = bookmark.Ancestors<Footer>().First();
                    footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(imageRun));
                }
                else
                {
                    // Default to main body
                    parent.InsertAfterSelf(imageRun);
                }
            }
        }

        // Validate image data
        private static bool ValidateImageData(byte[] imageData)
        {
            try
            {
                using (var ms = new MemoryStream(imageData))
                using (var img = System.Drawing.Image.FromStream(ms))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        // Insert a QR code at a bookmark
        private static void InsertQRCodeBookmark(WordprocessingDocument wordDoc, string bookmarkName, byte[] qrCodeBytes)
        {
            var bookmark = FindBookmark(wordDoc, bookmarkName);
            if (bookmark != null && ValidateImageData(qrCodeBytes))
            {
                var parent = bookmark.Parent;
                var qrImageRun = CreateImageRun(wordDoc, qrCodeBytes, "qrcode.png", 110, 60);

                // Check if the bookmark is in a header or footer
                if (bookmark.Ancestors<Header>().Any())
                {
                    var header = bookmark.Ancestors<Header>().First();
                    header.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(qrImageRun));
                }
                else if (bookmark.Ancestors<Footer>().Any())
                {
                    var footer = bookmark.Ancestors<Footer>().First();
                    footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(qrImageRun));
                }
                else
                {
                    // Default to main body
                    parent.InsertAfterSelf(qrImageRun);
                }
            }
        }

        // Create an image run for inserting
        private static DocumentFormat.OpenXml.Wordprocessing.Run CreateImageRun(
            WordprocessingDocument wordDoc,
            byte[] imageBytes,
            string imageFileName,
            int widthInPx,
            int heightInPx)
        {
            const int emuPerPixel = 9525;
            long widthEmu = widthInPx * emuPerPixel;
            long heightEmu = heightInPx * emuPerPixel;

            var imagePart = wordDoc.MainDocumentPart.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(stream);
            }

            var imageRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
            string imageId = wordDoc.MainDocumentPart.GetIdOfPart(imagePart);
            AddImageToRun(imageRun, imageId, widthEmu, heightEmu);
            return imageRun;
        }

        // Add image reference to a run
        private static void AddImageToRun(DocumentFormat.OpenXml.Wordprocessing.Run run, string imageId, long widthEmu, long heightEmu)
        {
            var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = widthEmu, Cy = heightEmu },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties
                    {
                        Id = (UInt32Value)1U,
                        Name = "Image"
                    },
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                    new DocumentFormat.OpenXml.Drawing.Blip { Embed = imageId },
                                    new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                                        new DocumentFormat.OpenXml.Drawing.Offset { X = 0L, Y = 0L },
                                        new DocumentFormat.OpenXml.Drawing.Extents { Cx = widthEmu, Cy = heightEmu })))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));

            run.AppendChild(drawing);
        }
        
        private static byte[] MakeQrCode(string text, int pixelsPerModule = 20)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(text))
                {
                    //throw new ArgumentException("Input text cannot be null or empty.", nameof(text));
                    text = "QRCode";
                }
                // Validate payload size
                if (text.Length > 1663) // Adjust based on encoding and ECC level
                {
                    //throw new ArgumentException("Payload exceeds the maximum size for the QR code.");
                    text = "QRCode";
                }

                using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
                {
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode(text, QRCodeGenerator.ECCLevel.Q);
                    using (QRCode qrCode = new QRCode(qrCodeData))
                    using (Bitmap qrCodeImage = qrCode.GetGraphic(pixelsPerModule))
                    {
                        // Resize the QR code image to 110x60 pixels
                        using (Bitmap resizedImage = new Bitmap(qrCodeImage, new System.Drawing.Size(110, 60)))
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            resizedImage.Save(memoryStream, ImageFormat.Png);
                            return memoryStream.ToArray();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public static byte[] GenerateQRCode(string qrData)
        {
            using (var qrGenerator = new QRCodeGenerator())
            {
                var qrCodeData = qrGenerator.CreateQrCode(qrData, QRCodeGenerator.ECCLevel.Q);
                using (var qrCode = new QRCode(qrCodeData))
                using (var bitmap = qrCode.GetGraphic(20)) // Adjust pixel size as needed
                using (var stream = new MemoryStream())
                {
                    bitmap.Save(stream, ImageFormat.Png);
                    return stream.ToArray();
                }
            }
        }


        public byte[] ImageToByteArray(System.Drawing.Image imageIn)
        {
            using (var ms = new MemoryStream())
            {
                imageIn.Save(ms, imageIn.RawFormat);
                return ms.ToArray();
            }
        }

        public static byte[] CompressPayload(string payload)
        {
            using (var memoryStream = new MemoryStream())
            using (var gzipStream = new GZipStream(memoryStream, CompressionMode.Compress))
            using (var writer = new StreamWriter(gzipStream))
            {
                writer.Write(payload);
                writer.Close();
                return memoryStream.ToArray();
            }
        }

        public static byte[] getSignFromDB(string userName)
        {
            string connectionString = "Server=localhost;Database=CenteralUserInfo;User Id=sa;Password=Aa@12345;";
            string query = "SELECT FIRST_SIGNATURE FROM CenteralUserInfo.dbo.Pargar_USER_SIGN WHERE USER_NAME = @UserName";
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

                    if (result is string hexString)
                    {
                        Console.WriteLine($"Hex string data retrieved for user: {userName}");

                        try
                        {
                            return HexStringToByteArray(hexString);
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

        public static byte[] HexStringToByteArray(string hexString)
        {
            if (string.IsNullOrWhiteSpace(hexString))
                throw new ArgumentException("Hex string cannot be null or empty.", nameof(hexString));

            // Ensure even-length string (each byte is 2 hex characters)
            if (hexString.Length % 2 != 0)
                throw new FormatException("Invalid hex string. Length must be a multiple of 2.");

            byte[] byteArray = new byte[hexString.Length / 2];
            for (int i = 0; i < byteArray.Length; i++)
            {
                string hexPair = hexString.Substring(i * 2, 2);
                byteArray[i] = Convert.ToByte(hexPair, 16);
            }

            return byteArray;
        }

        public static System.Drawing.Image BinaryToImage(byte[] binaryData)
        {
            if (binaryData == null || binaryData.Length == 0)
                throw new ArgumentException("Binary data cannot be null or empty.", nameof(binaryData));
            try
            {
                using (MemoryStream memoryStream = new MemoryStream(binaryData))
                {
                    return System.Drawing.Image.FromStream(memoryStream);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to convert binary data to an image.", ex);
            }
        }
        /*FindBookmark old
        public static BookmarkStart FindBookmark(WordprocessingDocument wordDoc, string bookmarkName)
        {
            // Search bookmarks in the main document body
            var bookmarks = wordDoc.MainDocumentPart.Document.Descendants<BookmarkStart>()
                .Where(b => b.Name == bookmarkName);

            if (bookmarks.Any())
                return bookmarks.First();

            // Search in headers and footers
            foreach (var header in wordDoc.MainDocumentPart.HeaderParts)
            {
                var headerBookmarks = header.RootElement.Descendants<BookmarkStart>()
                    .Where(b => b.Name == bookmarkName);
                if (headerBookmarks.Any())
                    return headerBookmarks.First();
            }

            foreach (var footer in wordDoc.MainDocumentPart.FooterParts)
            {
                var footerBookmarks = footer.RootElement.Descendants<BookmarkStart>()
                    .Where(b => b.Name == bookmarkName);
                if (footerBookmarks.Any())
                    return footerBookmarks.First();
            }

            // Search in footnotes, endnotes, and other parts if needed
            if (wordDoc.MainDocumentPart.FootnotesPart != null)
            {
                var footnoteBookmarks = wordDoc.MainDocumentPart.FootnotesPart.RootElement.Descendants<BookmarkStart>()
                    .Where(b => b.Name == bookmarkName);
                if (footnoteBookmarks.Any())
                    return footnoteBookmarks.First();
            }

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
        */
        // Find a bookmark by its name
        private static DocumentFormat.OpenXml.Wordprocessing.BookmarkStart FindBookmark(WordprocessingDocument wordDoc, string bookmarkName)
        {
            // Search bookmarks in the main document body
            var bookmark = wordDoc.MainDocumentPart.Document
                .Descendants<DocumentFormat.OpenXml.Wordprocessing.BookmarkStart>()
                .FirstOrDefault(b => b.Name == bookmarkName);

            if (bookmark != null)
                return bookmark;

            // Search in headers
            foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
            {
                bookmark = headerPart.RootElement
                    .Descendants<DocumentFormat.OpenXml.Wordprocessing.BookmarkStart>()
                    .FirstOrDefault(b => b.Name == bookmarkName);
                if (bookmark != null)
                    return bookmark;
            }

            // Search in footers
            foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
            {
                bookmark = footerPart.RootElement
                    .Descendants<DocumentFormat.OpenXml.Wordprocessing.BookmarkStart>()
                    .FirstOrDefault(b => b.Name == bookmarkName);
                if (bookmark != null)
                    return bookmark;
            }

            return null; // Bookmark not found
        }


        private static void InsertImageAtBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmark, string imagePath)
        {
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            var mainPart = wordDoc.MainDocumentPart;
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);

            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            const long widthEmu = 110 * 914400 / 96; // 914400 EMUs per inch, 96 DPI
            const long heightEmu = 60 * 914400 / 96;

            var drawingElement = CreateImageElement(mainPart.GetIdOfPart(imagePart), widthEmu, heightEmu);
            var imageRun = new DocumentFormat.OpenXml.Wordprocessing.Run(drawingElement);

            // Check if the bookmark is in the header/footer or main body
            var parent = bookmark.Parent;
            if (IsInHeaderOrFooter(bookmark))
            {
                InsertInHeaderOrFooter(wordDoc, parent, imageRun);
            }
            else
            {
                bookmark.Parent.InsertAfter(imageRun, bookmark);
            }
        }

        private static bool IsInHeaderOrFooter(BookmarkStart bookmark)
        {
            var parent = bookmark.Ancestors<Header>().FirstOrDefault();
            return parent != null || bookmark.Ancestors<Footer>().Any();
        }

        private static void InsertInHeaderOrFooter(WordprocessingDocument wordDoc, OpenXmlElement parent, DocumentFormat.OpenXml.Wordprocessing.Run imageRun)
        {
            if (parent.Ancestors<Header>().Any())
            {
                var header = parent.Ancestors<Header>().First();
                header.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(imageRun));
            }
            else if (parent.Ancestors<Footer>().Any())
            {
                var footer = parent.Ancestors<Footer>().First();
                footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(imageRun));
            }
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
                    new DocumentFormat.OpenXml.Drawing.NonVisualGraphicFrameDrawingProperties(
                        new GraphicFrameLocks { NoChangeAspect = true }),
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

        private static (long width, long height) GetImageDimensions(string imagePath, long maxWidthEmu = 3000000, long maxHeightEmu = 3000000)
        {
            using (var image = SixLabors.ImageSharp.Image.Load(imagePath))
            {
                const int emusPerInch = 914400; // Conversion factor

                var dpiX = image.Metadata.HorizontalResolution > 0 ? image.Metadata.HorizontalResolution : 96; // Default DPI
                var dpiY = image.Metadata.VerticalResolution > 0 ? image.Metadata.VerticalResolution : 96;

                var widthEmu = (long)(image.Width * emusPerInch / dpiX);
                var heightEmu = (long)(image.Height * emusPerInch / dpiY);

                // Scale to fit within max dimensions
                if (widthEmu > maxWidthEmu || heightEmu > maxHeightEmu)
                {
                    double widthScale = (double)maxWidthEmu / widthEmu;
                    double heightScale = (double)maxHeightEmu / heightEmu;
                    double scale = Math.Min(widthScale, heightScale);

                    widthEmu = (long)(widthEmu * scale);
                    heightEmu = (long)(heightEmu * scale);
                }

                return (widthEmu, heightEmu);
            }
        }
    
    }
}
