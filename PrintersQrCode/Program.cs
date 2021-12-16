using System;
using System.IO;
using System.Security.Claims;
using System.Security.Principal;
using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.Extensions.Configuration.UserSecrets;
using QRCoder;
using SelectPdf;

namespace PrintersQrCode
{
	class Program
	{
		static void Main(string[] args)
		{


			if (args.Length == 3)
			{
				
				var QrBytes = QrCodeGenerator(args[0], args[1], args[2]);
				//throw new Exception(QrBytes);
				if (SaveQrCodeAsPdf(QrBytes,args[2]))
					Console.WriteLine($"QrCode Generate Sucssefuly Go To Path To View");
			}
			else
			{
				throw new Exception("no args");
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
		public static string QrCodeGenerator(string DocumentNo , string DepartmentID,string UserName) 
		{
			try
			{
				QRCodeGenerator QrGenertator = new QRCodeGenerator();
				//string UserName = //Environment.UserName;
				string NewLine = Environment.NewLine;
				string Payload = string.Format("{0}{1}{2}{3}{4}",
								DocumentNo, NewLine, DepartmentID, NewLine, UserName);
				var QrData = QrGenertator.CreateQrCode(Payload,
											   QRCodeGenerator.ECCLevel.Q);
				PdfByteQRCode pdfBytesQrCode = new PdfByteQRCode(QrData);
				Base64QRCode Base = new Base64QRCode(QrData);
				
				return Base.GetGraphic(20);
			}
			catch (Exception ex)
			{
				throw ;
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
				throw ;
				string ErrorMsg = string.Format("Error Message: {0}\n stacktrace : {1}",
														ex.Message, ex.StackTrace);
				Console.WriteLine(ErrorMsg);
				return false;
			}
			
		}

		public static bool SaveQrCodeAsPdf(string bytes,string userName)
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
				string QrCodeImgTag = string.Format(HtmlBuilder.ToString(),imgSrc,QrHight,QrWidth);
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
		public static void Aspose(byte[] bitmapByteQR)
		{
			






		}
		// Register AppSetting.json configuration provider File at Configurations
		private static IConfiguration Configuration
		{
			get
			{
				return new ConfigurationBuilder()
				.SetBasePath(Directory.GetCurrentDirectory())
				.AddJsonFile("AppSettings.json", true, true)
				.Build();
			}
		}
	}
}
