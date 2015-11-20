using Excel2JsonConverter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sample
{
	class Program
	{
		static void Main(string[] args)
		{
			var table = ExcelTableReader.Read("Asset/sample.xlsx");

			using (FileStream stream = File.Open("Asset/sample.json", FileMode.Create, FileAccess.Write))
			{
				using (var writer = new JsonTableWriter(new StreamWriter(stream)))
				{
					writer.Write(table);
				}
			}

			Console.Out.WriteLine("Done.");
			Console.ReadKey();
		}
	}
}
