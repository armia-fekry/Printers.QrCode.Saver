using System;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Configuration;
using Spire.Pdf;
using Spire.Pdf.Annotations;
using Spire.Pdf.Annotations.Appearance;
using Spire.Pdf.Graphics;
using ZXing;
using ZXing.Common;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using System.Drawing;
using System.Collections.Generic;
using Microsoft.VisualBasic;

namespace PrintersQrCode
{

	class Program
    {
		public static string _Code { get; set; }
		public enum DepartmentType
        {
            Hr=07001,Finance=07002
        }
        public enum CodeType {
            Qrcode, BarCode
        }
        public static Dictionary<string,string> DepartmentsKeyValue = 
            new Dictionary<string, string>();

        public static void Main(string[] args)
        {
			try
			{
                DepartmentsKeyValue.Add("hr", "07001");
                DepartmentsKeyValue.Add("finance", "07002");
                generate(args);
            }
			catch (Exception ex)
			{
                Console.WriteLine(ex.Message);
				throw;
			}
           
        }
        private static bool generate(string[] args)
        {
            //autmoatic insert (hr user path)
            if (args.Count()==3)
            {
             
                string DocumentNo = AddNewToJson(args[0]).ToString();
                string depCode = DepartmentsKeyValue[args[0].ToLower()];
                var QrImage = QrCodeGenerator(depCode,DocumentNo,args[1]);
                var BarCodeImage = GenerateBarcode(depCode,DocumentNo);
                if (QrImage is null)
                    throw new Exception("Cannot Generate Qr Image");
                if (BarCodeImage is null)
                    throw new Exception("could not generte Barcode image"); 

                AddStampToPdf(args[2], QrImage);
                AddStampToPdf(args[2], BarCodeImage, CodeType.BarCode);
                return true;
            }
            if (args.Count() == 4)//manual insert
            {
                if (!UpdateExitingJson(args[1], args[0]))
                { throw new Exception("Somthing Wrong During Insert in Json, or Document No Not Found");  }
                string depCode = DepartmentsKeyValue[args[0].ToLower()];

                var QrImage = QrCodeGenerator(depCode, args[1], args[2]);
                var BarCodeImage = GenerateBarcode(depCode,args[1]);
                if (QrImage is null)
                    throw new Exception("Cannot Generate Qr Image");
                if (BarCodeImage is null)
                    throw new Exception("could not generte Barcode image");

                AddStampToPdf(args[3], QrImage);
                AddStampToPdf(args[3], BarCodeImage, CodeType.BarCode);
                return true;
            }
            else {
                return false;
                throw new Exception("Invalid Paramter Count");
            }
        }
        private static int AddNewToJson(string DepType)
        {
            string path = IntializeFile(DepType);
            int Value;
            var TextValue = File.ReadAllText(path);
            int.TryParse(TextValue, out Value);
            int lastId = Value == default(int) ? 500000000 : Value;
            File.WriteAllText(path, (lastId+1).ToString());
            return ++lastId;
        }

        private static bool UpdateExitingJson( string DocNo, string DepType)
        {
            string path = IntializeFile(DepType);
            Int64 LastValue;
            Int64.TryParse(File.ReadAllText(path), out LastValue);
            var DocumentNo = Convert.ToInt64(DocNo);
            if (DocumentNo > LastValue || DocumentNo<500000000 )
                return false;
            return true;
        }
        private static Image GenerateBarcode(string DepCode,string Sequence)
        {
            string _Data = $"007503{DepCode}{Sequence}";
            _Code = _Data;
            BarcodeWriter writer = new BarcodeWriter()
            {
                Format = BarcodeFormat.CODE_128,
                Options = new EncodingOptions
                {
                    Height = 65,
                    Width = 110,
                    PureBarcode = false,
                    Margin = 2
                },
            };
            return writer.Write(_Data);
        }

		private static string IntializeFile(string DepType)
		{
            string FilePathStr = $"{Configuration.GetSection("JsonFilePath").Value}\\";
            if (!Directory.Exists(FilePathStr))
                Directory.CreateDirectory(FilePathStr);

            if (DepType.ToLower() == "hr")
                FilePathStr = FilePathStr + "Hr.txt";
            if (DepType.ToLower() == "finance")
                FilePathStr = FilePathStr + "Finance.Txt";
            if (!File.Exists(FilePathStr))
                File.WriteAllText(FilePathStr, "");

            return FilePathStr;
			
		}

	    public static Image QrCodeGenerator(string DepCode,string DocumentNo,string UserName)
        {
            try
            {
                string newLine = Environment.NewLine;
                string Date = DateTime.Now.ToString("MM/dd/yyyy");
                string Time = DateTime.Now.ToShortTimeString();
                string MachineName = Environment.MachineName;
                string Department = DepartmentsKeyValue.FirstOrDefault(e => e.Value == DepCode).Key;
                string payload = $"007033{DepCode}{DocumentNo}{newLine}{Date} {Time}{newLine}Username: {UserName}{newLine}PC Name: {MachineName}{newLine}Department: {Department}";
                var _options = new QrCodeEncodingOptions()
                {
                    DisableECI = true,
                    CharacterSet = "UTF-8",
                    Width = 200,
                    Height = 200,
                    Margin=0,
                    ErrorCorrection= ErrorCorrectionLevel.H,
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
            PdfPageBase page = document.Pages[0];
            Spire.Pdf.Graphics.PdfTemplate template = new Spire.Pdf.Graphics.PdfTemplate(qrImage.Size);

            PdfImage image = PdfImage.FromImage(qrImage);
            int FontSize = 11;
            if (_codeType == CodeType.BarCode)
            {
                template.Graphics.DrawImage(image, 0, 0, image.Width, image.Height - 15);
                template.Graphics.DrawString($"{DateTime.Now.ToString("MM/dd/yy")} {DateTime.Now.ToShortTimeString()}",
                                   new Spire.Pdf.Graphics.PdfFont(PdfFontFamily.Helvetica, FontSize, PdfFontStyle.Bold),
                                   PdfBrushes.Black, new PointF(image.Width/6, image.Height - 15));
            }
            else
            {
                template.Graphics.DrawImage(image, 0, 0, image.Width, image.Height);         
            }
            RectangleF rectangle = new RectangleF(new PointF(0, 0), template.Size);
            PdfRubberStampAnnotation stamp = new PdfRubberStampAnnotation(rectangle);

            //set the appearance of the annotation  
            PdfAppearance appearance = new PdfAppearance(stamp) { Normal = template };
            stamp.Appearance = appearance;

            //add the stamp annotation to the page and save changes to file  
            page.AnnotationsWidget.Add(stamp);

            document.SaveToFile($@"{pdfPath}", FileFormat.PDF);

            if (_codeType == CodeType.BarCode)
            {
                var extention = Path.GetExtension(pdfPath);
                var directory = Path.GetDirectoryName(pdfPath);
                FileSystem.Rename(pdfPath, $"{directory}\\{_Code}{extention}");
            }

            return true;
        }

		#region Extra
		//private static Image ResizeImage(Image imgToResize, Size size)
		//{
		//    //Get the image current width  
		//    int sourceWidth = imgToResize.Width;
		//    //Get the image current height  
		//    int sourceHeight = imgToResize.Height;
		//    float nPercent = 0;
		//    float nPercentW = 0;
		//    float nPercentH = 0;
		//    //Calulate  width with new desired size  
		//    nPercentW = ((float)size.Width / (float)sourceWidth);
		//    //Calculate height with new desired size  
		//    nPercentH = ((float)size.Height / (float)sourceHeight);
		//    if (nPercentH < nPercentW)
		//        nPercent = nPercentH;
		//    else
		//        nPercent = nPercentW;
		//    //New Width  
		//    int destWidth = (int)(sourceWidth * nPercent);
		//    //New Height  
		//    int destHeight = (int)(sourceHeight * nPercent);
		//    Bitmap b = new Bitmap(destWidth, destHeight);
		//    Graphics g = Graphics.FromImage((System.Drawing.Image)b);
		//    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
		//    // Draw image with new width and height  
		//    g.DrawImage(imgToResize, 0, 0, destWidth, destHeight);
		//    g.Dispose();
		//    return b;
		//}

		//private static bool GenerateJsonFile(string FilePath) 
		//{
		//    if (File.Exists(FilePath))
		//        return false;
		//    return true;
		//} 
		#endregion

	}
}
