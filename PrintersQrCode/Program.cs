using System;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Configuration;
using Spire.Pdf;
using Spire.Pdf.Annotations;
using Spire.Pdf.Annotations.Appearance;
using Spire.Pdf.Graphics;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using System.Drawing;
using System.Collections.Generic;
using Microsoft.VisualBasic;
using NetBarcode;
using System.DirectoryServices;
using System.Collections;

namespace PrintersQrCode
{

	class Program
    {
		public static string _Code { get; set; }
		public static string _UserName { get; set; }
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
				// GetAllUsers();

				DepartmentsKeyValue.Add("hr", "07001");
				DepartmentsKeyValue.Add("finance", "07002");
				//args = new string[] { "finance", "armiafekryzaki", @"D:\armia\00750307002500000024.pdf" };
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

                _UserName = args[1];
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
                _UserName = args[2];
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

            var barcode = new Barcode(_Data,NetBarcode.Type.Code128,false,300,50);

			//BarcodeWriterGeneric writer = new BarcodeWriterGeneric()
			//{
			//	Format = BarcodeFormat.CODABAR,
			//	Options = new EncodingOptions
			//	{
			//		Height = 65,
			//		Width = 110,
			//		PureBarcode = false,
			//		Margin = 0,
			//		GS1Format = true
			//	}
			//};
			//return writer.Write(_Data);
			return barcode.GetImage();
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
        private static void GetAllUsers()
        {
            SearchResultCollection results;
            DirectorySearcher ds = null;
            DirectoryEntry de = new DirectoryEntry(GetCurrentDomainPath());
            
            ds = new DirectorySearcher(de);
            ds.PropertiesToLoad.Add("department");
            ds.PropertiesToLoad.Add("objectSid");
            ds.Filter = "(&((&(objectCategory=Person)(objectClass=User)))(samaccountname=" + "armia.fekry" + "))";
           
            results = ds.FindAll();

            if (results != null)
            {
                foreach (SearchResult result in results)
                {
                    foreach (DictionaryEntry property in result.Properties)
                    {
                        Console.Write(property.Key + ": ");
                        foreach (var val in (property.Value as ResultPropertyValueCollection))
                        {
                            Console.Write(val + "; ");
                        }
                        Console.WriteLine("");
                    }
                }
            }
        }

        private static string GetCurrentDomainPath()
        {
            DirectoryEntry de = new DirectoryEntry("LDAP://RootDSE");

            return "LDAP://" + de.Properties["defaultNamingContext"][0].ToString();
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
                    Width = 300,
                    Height = 300,
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
            var _size = _codeType == CodeType.BarCode ? new Size(230,80):new Size(100,100);
            Spire.Pdf.Graphics.PdfTemplate template = new Spire.Pdf.Graphics.PdfTemplate(_size);

            PdfImage image = PdfImage.FromImage(qrImage);
            int FontSize = 11;
            string CodeDate = $"{DateTime.Now.ToString("MM/dd/yy")} {DateTime.Now.ToShortTimeString()}";
           var CodeStringSize =  MeasureString(_Code);
           var CodeDateStringSize = MeasureString(CodeDate);

            if (_codeType == CodeType.BarCode)
            {
                template.Graphics.DrawImage(image,0, 0, _size.Width, _size.Height-40);
                template.Graphics.DrawString(_Code,
                    new Spire.Pdf.Graphics.PdfFont(PdfFontFamily.Helvetica, 12, PdfFontStyle.Regular),
                    PdfBrushes.Black,((template.Width- CodeStringSize.Width)/2)+20, 42)
                  ;

				template.Graphics.DrawString(CodeDate,
								   new Spire.Pdf.Graphics.PdfFont(PdfFontFamily.Helvetica, 12, PdfFontStyle.Bold),
								   PdfBrushes.Black, ((template.Width -CodeDateStringSize.Width)/ 2)+20, 65);
			}
            else
            {
                template.Graphics.DrawImage(image, 0, 0, _size.Width, _size.Height);         
            }
            RectangleF rectangle;
            if (_codeType == CodeType.Qrcode)
            { rectangle = new RectangleF(new PointF(495,700), template.Size); }
            else { rectangle = new RectangleF(new PointF(0, 720), template.Size); }
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
                FileSystem.Rename(pdfPath, $"{directory}\\{_UserName}#{_Code}{extention}");
            }

            return true;
        }
        private static SizeF MeasureString(string _Data) {
            Font f = new Font("Courier", 12, FontStyle.Regular);

            //create a bmp / graphic to use MeasureString on
            Bitmap b = new Bitmap(300, 300);
            Graphics g = Graphics.FromImage(b);

            //measure the string and choose a width:
            SizeF sizeOfString = new SizeF();
            sizeOfString = g.MeasureString(_Data, f,250);
            return sizeOfString;
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
