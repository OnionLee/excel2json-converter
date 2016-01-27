using Excel2JsonConverter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Sample
{
	class Program
	{
		static void Main(string[] args)
		{
			DirectoryInfo di = new DirectoryInfo(Directory.GetCurrentDirectory());
			var files = di.GetFiles("*.xlsx");

			foreach(var path in files)
			{
				Convert(path.FullName, path.Name);
			}

			Console.Out.WriteLine("Done");
			Console.ReadKey();
		}

		static void Convert(string path, string fileName)
		{
			try
			{
				var table = ExcelTableReader.Read(path);

				using (FileStream stream = File.Open(path.Replace("xlsx", "json"), FileMode.Create, FileAccess.Write))
				{
					using (var writer = new JsonTableWriter(new StreamWriter(stream)))
					{
						writer.Write(table);
						Console.Out.WriteLine(string.Format("[{0}]...OK", fileName));
					}
				}
			}
			catch(Exception e)
			{
				Console.Out.WriteLine(string.Format("[{0}]...Error : {1}", fileName, e.Message));
			}
			
		}
	}
}
