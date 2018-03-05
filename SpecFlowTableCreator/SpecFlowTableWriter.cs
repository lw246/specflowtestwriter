using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace SpecFlowTableCreator
{
    public class SpecFlowTableWriter
    {
        private DataSet _tablesFromFile;

        public static SpecFlowTableWriter New()
        {
            return new SpecFlowTableWriter();
        }

        public SpecFlowTableWriter ReadDataFromFile(string filePath)
        {
            _tablesFromFile = ExcelConnection.New()
                .ForFile(filePath)
                .GetAllTables();

            return this;
        }

        public Dictionary<string, string> Create()
        {
            Dictionary<string, string> specFlowTables = new Dictionary<string, string>();

            foreach (DataTable table in _tablesFromFile.Tables)
            {
                specFlowTables[SheetName(table.TableName)] = (CreateSpecFlowTable(table));
            }

            return specFlowTables;
        }

        private string SheetName(string tableName)
        {
            return tableName.Replace("$", "");
        }

        private string CreateSpecFlowTable(DataTable table)
        {
            var stringBuilder = new StringBuilder();

            stringBuilder.Append(CreateHeadings(table.Columns));
            
            foreach (DataRow row in table.Rows)
            {
                stringBuilder.AppendLine(CreateRow(row, table.Columns.Count));
            }

            return stringBuilder.ToString();
        }

        private string CreateRow(DataRow row, int numberOfColumns)
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.Append("|");

            for (int i=0; i<numberOfColumns; i++)
            {
                if (row.IsNull(i))
                {
                    stringBuilder.Append("<NULL>|");
                }
                else
                {
                    stringBuilder.Append($"{row[i]}|");
                }
            }

            return stringBuilder.ToString();
            
        }

        private string CreateHeadings(DataColumnCollection columns)
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.Append("|");

            foreach (DataColumn column in columns)
            {
                stringBuilder.Append($"{column.ColumnName}|");
            }
            return stringBuilder.ToString();
        }
    }
    
}
