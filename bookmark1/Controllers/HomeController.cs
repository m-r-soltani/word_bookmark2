using System.Diagnostics;
using bookmark1.Models;
using BookMarks;
using Microsoft.AspNetCore.Mvc;
using System.Globalization;

namespace bookmark1.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {


            //////////////////////////////////////text,image,qrcode
            //var bookmarksContent = new Dictionary<string, string>
            //{
            //    { "تصویر_واحد_سازمانی",  @"C:\mydocs\img\1.png" },
            //    { "واحد_سازمانی", "بازرسی کل استان تست" },
            //    { "تاریخ_خورشیدی",  DateTime.Now.ToString("dd-MM-yyyy") },
            //    { "پیوست", "ندارد"  },
            //    { "گیرندگان_رونوشت","تستی تستیان"  },
            //    { "رونوشت","تستی تستیان"  },
            //    { "شماره_ثبت","127126"  },
            //    { "طبقه_بندی","غیر محرمانه"  },
            //    { "عنوان_محترمانه_کامل_گیرندگان_رونوشت","تستی تستیان"  },
            //    { "فوریت", "فوری" },
            //    { "نام_و_نام_خانوادگی_فرستنده","تستی تستیان"  },
            //    { "نوع_جایگاه_امضاکننده_اصلی","تستی تستیان"  },
            //    { "آدرس_جایگاه_فرستنده", "تهران - خیابان طالقانی سازمان بازرسی کشور"  },
            //    { "اهمیت_", "مهم"  },
            //    { "بارکد_شمس", "QRCode"  },
            //    { "امضای_اصلی", "@Binary:6930"  },
            //    { "گیرندگان_اصلی","جناب آقای تستی تستیان \r\n رییس محترم شورای اسلامی شهر"  },
            //    { "نوع_جایگاه_امضاکننده_اصلی", "تستی تستیان"  },
            //};
            //BookmarkOpenxml.UpdateBookmarks(docPath, bookmarksContent);




            string docPath = @"C:\mydocs\f1.docx";
            //////////////////////////////////////simple text//////////////////////////////////////////

            var bmText = new Dictionary<string, string>
            {
                { "واحد_سازمانی", "بازرسی کل استان تست" },
                { "تاریخ_خورشیدی",  DateTime.Now.ToString("dd-MM-yyyy") },
                { "پیوست", "ندارد" },
                { "گیرندگان_رونوشت", "تستی تستیان" },
                { "رونوشت", "تستی تستیان"  },
                { "شماره_ثبت", "127126"  },
                { "طبقه_بندی", "غیر محرمانه"  },
                { "عنوان_محترمانه_کامل_گیرندگان_رونوشت", "تستی تستیان"  },
                { "فوریت", "فوری" },
                { "نام_و_نام_خانوادگی_فرستنده", "تستی تستیان"  },
                { "نوع_جایگاه_امضاکننده_اصلی", "تستی تستیان"  },
                { "آدرس_جایگاه_فرستنده", "تهران - خیابان طالقانی سازمان بازرسی کشور"  },
                { "اهمیت_", "مهم"  },
                { "گیرندگان_اصلی", "جناب آقای تستی تستیان \r\n رییس محترم شورای اسلامی شهر"  },
            };
            BMH.BMH.UpdateTextBookmarks(docPath, bmText);
            //////////////////////////////////////Binary Binary Image//////////////////////////////////////////
            var bmImage = new Dictionary<string, string>
            {
                { "تصویر_واحد_سازمانی", @"C:\mydocs\img\1.png" },
            };
            BMH.BMH.UpdateImageBookmarks(docPath, bmImage);
            //////////////////////////////////////Binary Binary Image//////////////////////////////////////////
            var bm_Binary = new Dictionary<string, string>
            {
                { "امضای_اصلی", "6936" },
            };
            BMH.BMH.UpdateBinaryImageBookmarks(docPath, bm_Binary);
            //////////////////////////////////////QrCode//////////////////////////////////////////
            var bm_QrCode = new Dictionary<string, string>
            {
                { "بارکد_شمس", "test" },
            };
            BMH.BMH.UpdateQrcodeBookmarks(docPath, bm_QrCode);
            //////////////////////////////////////Docx To Pdf//////////////////////////////////////////
            //OfficeHandler.WordHandler.ConvertWordToPdfWithLibreOffice(@"C:\mydocs\f1.docx", @"C:\mydocs\pdf\aaa.pdf");
            //OfficeHandler.WordHandler.ConvertWordToPdf();
            //OfficeHandler.WordHandler.docxtopdfapose();
            //OfficeHandler.WordHandler.ConvertWordToPdf();
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
