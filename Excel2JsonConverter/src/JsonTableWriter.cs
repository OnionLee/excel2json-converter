using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2JsonConverter
{
	public class JsonTableWriter : IDisposable
	{
		JsonSerializer serializer;
		JsonTextWriter writer;

		public JsonTableWriter(TextWriter textWriter)
		{
			serializer = new JsonSerializer();
			writer = new JsonTextWriter(textWriter);
			writer.Formatting = Formatting.Indented;
		}

		public void Write(Table table, string parentId = "")
		{
			writer.WriteStartObject();
			writer.WritePropertyName(table.TableName);
			writer.WriteStartArray();
			for (int i = 0; i<table.Rows.Count; i++)
			{
				var dic = table.Rows[i].AsDictionary();
				serializer.Serialize(writer, dic);
			}
			writer.WriteEndArray();
			writer.WriteEndObject();
		}

		public void Close()
		{
			Dispose();
		}

		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		protected virtual void Dispose(bool disposing)
		{
			if (disposing)
			{
				writer.Close();
			}
		}
	}
}
