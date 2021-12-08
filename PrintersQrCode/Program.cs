using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Security.Principal;
using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.Extensions.Configuration.UserSecrets;
using QRCoder;
using SelectPdf;
using Spire.Pdf;
using Spire.Pdf.Annotations;
using Spire.Pdf.Annotations.Appearance;
using Spire.Pdf.Graphics;
using PdfDocument = SelectPdf.PdfDocument;
using PdfTemplate = SelectPdf.PdfTemplate;

namespace PrintersQrCode
{
    class Program
    {
        static void Main(string[] args)
        {


            if (args.Length == 4)
            {

                var qrBytes = QrCodeGenerator(args[0], args[1], args[2]);
                if (!string.IsNullOrEmpty(qrBytes))
                {
                    try
                    {
                        var img = LoadImage(qrBytes);
                        AddStampToPdf(args[3], img);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }

                }
                //throw new Exception(QrBytes);
                //if (SaveQrCodeAsPdf(qrBytes, args[2]))
                //    Console.WriteLine($"QrCode Generate Sucssefuly Go To Path To View");
            }
            else
            {
                throw new Exception("no args " + args.Length  );
                Console.WriteLine("enter Valid Data and try again!");
            }
            //Console.ReadLine();
        }

        /// <summary>
        /// Function To Generte Qr Code As An Array Of Bytes
        /// </summary>
        /// <param name="DocumentNo"></param>
        /// <param name="DepartmentID"></param>
        /// <returns>Qrcode As Base24StringImg</returns>
        public static string QrCodeGenerator(string DocumentNo, string DepartmentID, string UserName)
        {
            try
            {
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                //string UserName = //Environment.UserName;
                string newLine = Environment.NewLine;
                string payload = $"{DocumentNo}{newLine}{DepartmentID}{newLine}{UserName}";
                var qrData = qrGenerator.CreateQrCode(payload,
                                               QRCodeGenerator.ECCLevel.Q);
                PdfByteQRCode pdfBytesQrCode = new PdfByteQRCode(qrData);
                Base64QRCode Base = new Base64QRCode(qrData);
                return Base.GetGraphic(20);
            }
            catch (Exception ex)
            {
                throw;
                string ErrorMsg = string.Format("Error Message: {0}\n stacktrace : {1}",
                                        ex.Message, ex.StackTrace);
                Console.WriteLine(ErrorMsg);
                return null;
            }
        }

        /// <summary>
        /// Method Will Take QrCode Array Of Bytes , Save it As Pdf
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns> True If Saved Successfully </returns>
        public static bool SaveQrCodeAsPdf(byte[] bytes)
        {
            try
            {
                string UserPath = Configuration.GetSection("Path")?.Value;
                string UserFileName = Configuration.GetSection("FileName")?.Value;
                string UserName = Environment.UserName;

                //path = UserPath\UserName
                string Path = String.Format("{0}\\{1}", UserPath, UserName);

                if (!Directory.Exists(Path))
                    Directory.CreateDirectory(Path);

                //path = UserPath\UserName\UserFileName.pdf
                string QrPdfFile = string.Format("{0}\\{1}.pdf", Path, UserFileName);

                FileStream FileStream = new FileStream(
                    QrPdfFile, FileMode.Create, FileAccess.Write);
                BinaryWriter binaryWriter = new BinaryWriter(FileStream);
                binaryWriter.Write(bytes);
                binaryWriter.Close();
                FileStream.Close();
                return true;
            }
            catch (Exception ex)
            {
                throw;
                string ErrorMsg = string.Format("Error Message: {0}\n stacktrace : {1}",
                                                        ex.Message, ex.StackTrace);
                Console.WriteLine(ErrorMsg);
                return false;
            }

        }

        public static bool SaveQrCodeAsPdf(string bytes, string userName)
        {
            string UserPath = "C:\\"; //Configuration.GetSection("Path")?.Value;
                                      //	throw new Exception(UserPath);
            string UserFileName = "Test101"; //Configuration.GetSection("FileName")?.Value;
            string QrWidth = "80px"; //Configuration.GetSection("Width")?.Value;
            string QrHight = "80px";// Configuration.GetSection("Hight")?.Value;


            //throw new Exception("user name is :"+UserName);
            //path = UserPath\UserName
            //string Path = String.Format("{0}\\{1}", UserPath, UserName);
            //if(string.IsNullOrEmpty(Path))
            //throw new Exception("no path defined");
            //throw new Exception("Faddo");

            string UserName = userName;

            //path = UserPath\UserName
            string Path = String.Format("{0}\\{1}", UserPath, UserName);

            if (!Directory.Exists(Path))
                Directory.CreateDirectory(Path);

            //path = UserPath\UserName\UserFileName.pdf
            string QrPdfFile = string.Format("{0}\\{1}.pdf", Path, UserFileName);

            var imgSrc = String.Format("data:image/jpg;base64,{0}", bytes);

            StringBuilder HtmlBuilder = new StringBuilder();
            HtmlBuilder.Append(@"<div>");
            HtmlBuilder.Append("<img src ={0} width= {1} height = {2} />");
            HtmlBuilder.Append("</div>");
            string QrCodeImgTag = string.Format(HtmlBuilder.ToString(), imgSrc, QrHight, QrWidth);
            HtmlToPdf converter = new HtmlToPdf();
            var PdfDocument = converter.ConvertHtmlString(QrCodeImgTag);
            PdfDocument.Save(QrPdfFile);
            return true;


            //throw ;
            //string ErrorMsg = string.Format("Error Message: {0}\n stacktrace : {1}",
            //										ex.Message, ex.StackTrace);
            //Console.WriteLine(ErrorMsg);
            //return false;


        }

        // Register AppSetting.json configuration provider File at Configurations
        private static IConfiguration Configuration =>
            new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("AppSettings.json", true, true)
                .Build();

        //Convert Base64 to Image
        private static Image LoadImage(string base64)
        {
            byte[] bytes = Convert.FromBase64String(base64);
            using MemoryStream ms = new MemoryStream(bytes);
            return Image.FromStream(ms);
        }

        private static bool AddStampToPdf(string pdfPath, Image qrImage)
        {
            Spire.Pdf.PdfDocument document = new Spire.Pdf.PdfDocument($@"{pdfPath}");
            if (document.Pages.Count <= 0)
                return false;

            PdfPageBase page = document.Pages[0];

            //initialize a new PdfTemplate object  
            Spire.Pdf.Graphics.PdfTemplate template = new Spire.Pdf.Graphics.PdfTemplate(100, 100);
            //load an image, resize and draw it on template  
            var imageResize = qrImage;
            if (imageResize.Width > 100 && imageResize.Height > 100)
            {
                imageResize = ResizeImage(qrImage, new Size(100, 100));
            }

            PdfImage image = PdfImage.FromImage(imageResize);
            //float width = image.Width * 0.4f;
            //float height = image.Height * 0.4f;
            template.Graphics.DrawImage(image, 0, 0, image.Width, image.Height);
            //initialize an instance of PdfRubberStamoAnnotation class based on the size and position of RectangleF  
            RectangleF rectangle = new RectangleF(new PointF(20, 0), template.Size);
            PdfRubberStampAnnotation stamp = new PdfRubberStampAnnotation(rectangle);

            //set the appearance of the annotation  
            PdfAppearance appearance = new PdfAppearance(stamp) { Normal = template };
            stamp.Appearance = appearance;

            //add the stamp annotation to the page and save changes to file  
            page.AnnotationsWidget.Add(stamp);

            document.SaveToFile($@"{pdfPath}", FileFormat.PDF);
            return true;
        }
        private static Image ResizeImage(Image imgToResize, Size size)
        {
            //Get the image current width  
            int sourceWidth = imgToResize.Width;
            //Get the image current height  
            int sourceHeight = imgToResize.Height;
            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;
            //Calulate  width with new desired size  
            nPercentW = ((float)size.Width / (float)sourceWidth);
            //Calculate height with new desired size  
            nPercentH = ((float)size.Height / (float)sourceHeight);
            if (nPercentH < nPercentW)
                nPercent = nPercentH;
            else
                nPercent = nPercentW;
            //New Width  
            int destWidth = (int)(sourceWidth * nPercent);
            //New Height  
            int destHeight = (int)(sourceHeight * nPercent);
            Bitmap b = new Bitmap(destWidth, destHeight);
            Graphics g = Graphics.FromImage((System.Drawing.Image)b);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            // Draw image with new width and height  
            g.DrawImage(imgToResize, 0, 0, destWidth, destHeight);
            g.Dispose();
            return b;
        }
    }
}
