using System;
using System.IO;
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
			
			if (args.Length == 2)
			{
				var QrBytes = QrCodeGenerator(args[0], args[1]);
				if (SaveQrCodeAsPdf(QrBytes))
					Console.WriteLine($"QrCode Generate Sucssefuly Go To Path To View");
			}
			else
			{
				Console.WriteLine("Invalid Data , You Should Enter DocumentNo and Departmentno Separated by Space");
			}
			Console.ReadLine();
		}

		/// <summary>
		/// Function To Generte Qr Code As Base24StringImg
		/// </summary>
		/// <param name="DocumentNo"></param>
		/// <param name="DepartmentID"></param>
		/// <returns>Qrcode As Base24StringImg</returns>
		public static string QrCodeGenerator(string DocumentNo , string DepartmentID) 
		{
			try
			{
				QRCodeGenerator QrGenertator = new QRCodeGenerator();
				string UserName = Environment.UserName;
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
