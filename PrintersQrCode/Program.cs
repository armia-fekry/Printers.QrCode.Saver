using System;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SelectPdf;
using Spire.Pdf;
using Spire.Pdf.Annotations;
using Spire.Pdf.Annotations.Appearance;
using Spire.Pdf.Graphics;
using PdfDocument = SelectPdf.PdfDocument;
using ZXing;
using ZXing.Common;
using ZXing.QrCode;
using ZXing.QrCode.Internal;

namespace PrintersQrCode
{

    class Program
    {
        public enum CodeType {
            Qrcode, BarCode
        }
        public static int pauload = 0;
        public static void Main(string[] args)
        {
            generate(args);
        }
        private static void generate(string[] args)
        {

            if (args[0].ToLower() == "a")//autmoatic insert
            {
                if (args.Count() != 4)
                    Console.WriteLine("Invalid Arg Number You Select Automatic and you should Insert DepType(f or h) , DepNo,user,Path");
                string DocumentNo = AddNewToJson(args[1]).ToString();
                var QrImage = QrCodeGenerator(DocumentNo, args[1], args[2]);
                string payload = $"{DocumentNo}-{args[1]}-{args[2]}";
                pauload = payload.Count();
                var BarCodeImage = GenerateBarcode(payload);
                if (QrImage is null)
                    throw new Exception("Cannot Generate Qr Image");
                if (BarCodeImage is null)
                    throw new Exception("could not generte Barcode image"); 

                AddStampToPdf(args[3], QrImage);
                AddStampToPdf(args[3], BarCodeImage, CodeType.BarCode);
            }
            else//manual insert
            {

                if (args.Count() != 5)
                    Console.WriteLine("Invalid Arg Number You Select manual and you should Insert DepType(f or h) , DepNo,DepNo,user,Path");
                if (!UpdateExitingJson(args[1], args[2]))
                { Console.WriteLine("Somthing Wrong During Insert in Json, or Document No Not Found"); return; }

                var QrImage = QrCodeGenerator(args[1], args[2], args[3]);
                string payload = $"{args[1]}-{args[2]}-{args[3]}";
                var BarCodeImage = GenerateBarcode(payload);
                if (QrImage is null)
                    throw new Exception("Cannot Generate Qr Image");
                if (BarCodeImage is null)
                    throw new Exception("could not generte Barcode image");
               
                AddStampToPdf(args[4], QrImage);
                AddStampToPdf(args[4], BarCodeImage, CodeType.BarCode);
            }
        }
        private static int AddNewToJson(string DepType)
        {
            string path = IntializeJsonFile(DepType);
            string json = File.ReadAllText(path);
            var jsonObject = JObject.Parse(json);
            JArray values = jsonObject.GetValue("Values") as JArray;
            int lastId = values.LastOrDefault() == default(JToken) ? 0 : values.LastOrDefault().Value<int>();
            values.Add((lastId + 1).ToString());
            jsonObject["Values"] = values;
            string final = JsonConvert.SerializeObject(jsonObject, Formatting.Indented);
            File.WriteAllText(path, final);
            return ++lastId;
        }

        private static bool UpdateExitingJson( string DocNo, string DepType)
        {
            string path = IntializeJsonFile(DepType);
            string json = File.ReadAllText(path);
            var jsonObject = JObject.Parse(json);
            JArray values = jsonObject.GetValue("Values") as JArray;
            var x = values.Any(e => e.Value<int>() == Convert.ToInt32(DocNo));
            if (!x)
                return false;
            return true;
        }
        private static Image GenerateBarcode(string _data)
        {
            BarcodeWriter writer = new BarcodeWriter()
            {
                Format = BarcodeFormat.CODE_128,
                Options = new EncodingOptions
                {
                    Height = 60,
                    Width = 110,
                    PureBarcode = false,
                    Margin = 2
                },
            };
            return writer.Write(_data);
        }

		private static string IntializeJsonFile(string DepType)
		{
            string JsonFileStr = $"{Configuration.GetSection("JsonFilePath").Value}\\";
            if (!Directory.Exists(JsonFileStr))
                Directory.CreateDirectory(JsonFileStr);

            if (DepType.ToLower() == "hr")
                JsonFileStr = JsonFileStr + "hr.json";
            if (DepType.ToLower() == "finance")
                JsonFileStr = JsonFileStr + "Finance.json";
           if(!File.Exists(JsonFileStr))
                File.WriteAllText(JsonFileStr, @"{""Values"":[]}");

            return JsonFileStr;
			
		}

	    public static Image QrCodeGenerator(string DocumentNo,string DepartmentID, string UserName)
        {
            try
            {
                string newLine = Environment.NewLine;
                string payload = $"{DocumentNo}{newLine}{DepartmentID}{newLine}{UserName}";
                var _options = new QrCodeEncodingOptions()
                {
                    DisableECI = true,
                    CharacterSet = "UTF-8",
                    Width = 100,
                    Height = 100,
                    Margin=2,
                    ErrorCorrection= ErrorCorrectionLevel.H,
                    PureBarcode=true
                };
                BarcodeWriter writer = new BarcodeWriter
                {
                    Format = BarcodeFormat.QR_CODE,Options=_options
                };
                Bitmap aztecBitmap;
                var result = writer.Write(payload);
                return new Bitmap(result);           
            }
            catch (Exception ex)
            {
                string ErrorMsg = string.Format("Error Message: {0}\n stacktrace : {1}",
                                        ex.Message, ex.StackTrace);
                Console.WriteLine(ErrorMsg);
                return null;
            }
        }


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
                string ErrorMsg = string.Format("Error Message: {0}\n stacktrace : {1}",
                                                        ex.Message, ex.StackTrace);
                Console.WriteLine(ErrorMsg);
                return false;
            }

        }

		public static bool SaveQrCodeAsPdf(string bytes)
		{
			try
			{
				string UserPath = Configuration.GetSection("Path")?.Value;
				string UserFileName = Configuration.GetSection("FileName")?.Value;
				string QrWidth = Configuration.GetSection("Width")?.Value;
				string QrHight = Configuration.GetSection("Hight")?.Value;

				string UserName = Environment.UserName;

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
				string QrCodeImgTag = string.Format(HtmlBuilder.ToString(),imgSrc,QrHight,QrWidth);
				HtmlToPdf converter = new HtmlToPdf();
				var PdfDocument = converter.ConvertHtmlString(QrCodeImgTag);
				PdfDocument.Save(QrPdfFile);			
				return true;
			}
			catch (Exception ex)
			{
				string ErrorMsg = string.Format("Error Message: {0}\n stacktrace : {1}",
														ex.Message, ex.StackTrace);
				Console.WriteLine(ErrorMsg);
				return false;
			}

        }

        // Register AppSetting.json configuration provider File at Configurations
        private static IConfiguration Configuration =>
            new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("AppSettings.json", true, true)
                .Build();

        ////Convert Base64 to Image
        //private static Image LoadImage(string base64)
        //{
        //    byte[] bytes = Convert.FromBase64String(base64);
        //    using MemoryStream ms = new MemoryStream(bytes);
        //    return Image.FromStream(ms);
        //}

        private static bool AddStampToPdf(string pdfPath, Image qrImage, CodeType _codeType = CodeType.Qrcode)
        {
            if (!File.Exists($@"{pdfPath}"))
            {
                Spire.Pdf.PdfDocument document2 = new Spire.Pdf.PdfDocument();
                document2.Pages.Add();
                document2.SaveToFile($@"{pdfPath}");
            }
           
            Spire.Pdf.PdfDocument document = new Spire.Pdf.PdfDocument($@"{pdfPath}");
            if (document.Pages.Count <= 0)
                return false;
           // Size _size = _codeType == CodeType.BarCode ? (qrImage.Width > 110 || qrImage.Width < 80 ? new Size(110, 60) : qrImage.Size) : qrImage.Size;
            PdfPageBase page = document.Pages[0];
            Spire.Pdf.Graphics.PdfTemplate template = new Spire.Pdf.Graphics.PdfTemplate(qrImage.Size);

            PdfImage image = PdfImage.FromImage(qrImage);
            int FontSize = pauload < 11 ? 6 : 9;
            if (_codeType == CodeType.BarCode)
            {
                template.Graphics.DrawImage(image, 0, 0, image.Width, image.Height - 10);
                template.Graphics.DrawString($"Date:{DateTime.Now.ToShortDateString()}     Time:{DateTime.Now.ToShortTimeString()}",
                                   new Spire.Pdf.Graphics.PdfFont(PdfFontFamily.Helvetica, FontSize, PdfFontStyle.Regular),
                                   PdfBrushes.Black, new PointF(10, image.Height - 10));
            }
            else
            {
                template.Graphics.DrawImage(image, 0, 0, image.Width, image.Height);         
            }
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

        private static bool GenerateJsonFile(string FilePath) 
        {
            if (File.Exists(FilePath))
                return false;
            return true;
        }
      
    }
}
