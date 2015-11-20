using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using System.IO;
using System.Data;

namespace Excel2JsonConverter
{
	public static class ExcelTableReader
	{
		private const int metadataIndex = 0;
		private const int columnNamesIndex = 1;
		private const int dataRowsStartIndex = 2;

		public static Table Read(string path)
		{
			var reader = GetExcelReader(path);
			var dataSet = reader.AsDataSet();

			var tables = new List<Table>(dataSet.Tables.Count);

			foreach (DataTable table in dataSet.Tables)
			{
				tables.Add(ConvertTable(table));
			}

			Postprocess(tables);

			return tables.FirstOrDefault(t => t.IsRootTable);
		}

		private static IExcelDataReader GetExcelReader(string path)
		{
			try
			{
				FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

				IExcelDataReader reader = null;

				if (path.EndsWith(".xls"))
				{
					reader = ExcelReaderFactory.CreateBinaryReader(stream);
				}
				if (path.EndsWith(".xlsx"))
				{
					reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
				}
				return reader;
			}
			catch (Exception)
			{
				throw;
			}
		}

		private static Table ConvertTable(DataTable table)
		{
			var tableName = table.TableName;
			var metaDatas = GetMetaDatas(table);
			var columnNames = GetColumnNames(table);
			var dataRows = GetDataRows(table, columnNames.Count);

			return new Table(tableName, metaDatas, columnNames, dataRows);
		}

		private static List<string> GetMetaDatas(DataTable table)
		{
			return GetRowAsStringList(table, 0);
		}

		private static List<string> GetColumnNames(DataTable table)
		{
			return GetRowAsStringList(table, 1);
		}

		private static List<string> GetRowAsStringList(DataTable table, int index)
		{
			try 
			{	        
				var row = table.Rows[index];
				var columnCount = table.Columns.Count;
				var result = new List<string>(columnCount);

				for (int i = 0; i < columnCount; i++)
				{
					result.Add(row[i].ToString()); 
				}

				return result;
			}
			catch (IndexOutOfRangeException)
			{
		
				throw;
			}
		}

		private static List<List<object>> GetDataRows(DataTable table, int columnCount)
		{
			try
			{
				var rowCount = table.Rows.Count - dataRowsStartIndex;
				var result = new List<List<object>>(rowCount);

				for (int i = dataRowsStartIndex; i < table.Rows.Count; i++)
				{
					columnCount = Math.Min(columnCount, table.Columns.Count);
					var dataList = new List<object>(columnCount);

					var row = table.Rows[i];
					for (int j = 0; j < columnCount; j++)
					{
						dataList.Add(row[j]);
					}
					result.Add(dataList);
				}

				return result;
			}
			catch (IndexOutOfRangeException)
			{
				
				throw;
			}
		}

		private static void Postprocess(List<Table> tables)
		{
			foreach (var table in tables)
			{
				if (table.HasParent)
				{
					var parentTable = tables.FirstOrDefault(t => table.ParentTableName == t.TableName);
					if(parentTable == null)
					{
						throw new Exception("ParentTable not found");
					}

					parentTable.AddChild(table);
				}
			}
		}
	}
}
