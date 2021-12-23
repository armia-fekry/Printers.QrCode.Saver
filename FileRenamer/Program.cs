using System;
using System.IO;
using System.Linq;
using Microsoft.VisualBasic.FileIO;

namespace FileRenamer
{
	class Program
	{
		static void Main(string[] args)
		{
			try
			{
				string FileName = GetFileName(args[0]);
				string Extension = Path.GetExtension(args[1]);

				var List = FileName.Split(',');
				var NewName = List?.FirstOrDefault(e => e.StartsWith("007033"));
				FileSystem.RenameFile(args[1], $"{NewName}{Extension}");
				File.Delete(args[0]);

			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				
			}
		
		}
		public static string  GetFileName(string FileName)
		{
			if (!File.Exists(FileName))
				throw new Exception($"File Not Exist With This Path {FileName}");
			return Path.GetFileNameWithoutExtension(FileName);
		}
	}
}
