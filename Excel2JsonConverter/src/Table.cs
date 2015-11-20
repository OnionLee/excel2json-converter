using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2JsonConverter
{
	public struct Row
	{
		private Table table;

		private List<object> field;

		public Row(Table table, List<object> field)
		{
			this.field = field;
			this.table = table;
		}

		public Dictionary<string, object> AsDictionary()
		{
			var dic = new Dictionary<string, object>(field.Count);

			for (int i = 0; i < field.Count; i++)
			{
				dic.Add(table.GetColumnName(i), field[i]);
			}

			if (table.HasChild)
			{
				foreach (var t in table.ChildTables)
				{
					var childRows = t.Rows.Where(r => r.AsDictionary()["parentId"] == dic["id"]);
					var dicList = childRows.Select(c => c.AsDictionary()).ToList();
					dic.Add(t.TableName, dicList);
				}
			}

			return dic;
		}
	}

	public class Table
	{
		private const string rootTableKey = "*root";

		private const int parentTableNameIndex = 0;

		private string tableName;
		private string parentTableName;

		private List<string> metaDatas;
		private List<string> columnNames;

		private List<Row> rows;

		private List<Table> childTables;

		public string TableName { get { return tableName; } }

		public bool IsRootTable { get { return parentTableName == rootTableKey; } }
		public bool HasParent { get { return !IsRootTable; } }
		public bool HasChild { get { return childTables != null; } }
		
		public string ParentTableName { get { return parentTableName; } }

		public IList<Row> Rows { get { return rows.AsReadOnly(); } }
		public IList<Table> ChildTables { get { return childTables.AsReadOnly(); } }

		public Table(string tableName, List<string> metaDatas, List<string> columnNames, List<List<object>> dataRows)
		{
			childTables = new List<Table>();

			this.tableName = tableName;
			this.metaDatas = metaDatas;
			this.columnNames = columnNames;
			this.rows = GetRows(dataRows);
			this.parentTableName = GetParentTableName();

		}

		private string GetParentTableName()
		{
			try
			{
				var name = metaDatas[parentTableNameIndex];
				if (string.IsNullOrEmpty(name))
				{
					throw new Exception("ParentTableName is empty");
				}

				return name;
			}
			catch (Exception)
			{
				throw;
			}
		}

		private List<Row> GetRows(List<List<object>> dataRows)
		{
			var result = new List<Row>(dataRows.Count);
			foreach (var dataRow in dataRows)
			{
				var row = new Row(this, dataRow);
				result.Add(row);
			}

			return result;
		}

		public string GetColumnName(int index)
		{
			try
			{
				return columnNames[index];
			}
			catch (IndexOutOfRangeException e)
			{
				throw;
			}
		}

		public void AddChild(Table child)
		{
			if (childTables == null)
				childTables = new List<Table>();

			childTables.Add(child);
		}
	}
}
