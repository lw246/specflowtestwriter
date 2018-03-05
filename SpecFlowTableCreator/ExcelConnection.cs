using System.Data.OleDb;
using System.Data;
using System.Collections.Generic;
using System.Text;
using SpecFlowTableCreator.interfaces;

namespace SpecFlowTableCreator
{
    public class ExcelConnection : FileConnection
    {
        private OleDbConnection _connection;

        public static ExcelConnection New()
        {
            return new ExcelConnection();
        }

        public FileConnection ForFile(string filePath)
        {
            _connection = new OleDbConnection(GetConnectionString(filePath));
            return this;
        }

        private string GetConnectionString(string filePath)
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();

            properties["Provider"] = "Microsoft.ACE.OLEDB.12.0";
            properties["Extended Properties"] = "Excel 12.0 XML";
            properties["Data Source"] = filePath;

            var stringBuilder = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in properties)
            {
                stringBuilder.Append(prop.Key);
                stringBuilder.Append("=");
                stringBuilder.Append(prop.Value);
                stringBuilder.Append(";");
            }

            return stringBuilder.ToString();
        }

        public DataSet GetAllTables()
        {
             var dataSet = new DataSet();

            using (_connection)
            {
                _connection.Open();
                var command = new OleDbCommand();
                command.Connection = _connection;

                var sheetTable = _connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                foreach (DataRow row in sheetTable.Rows)
                {
                    string sheetName = row["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    command.CommandText = $"SELECT * FROM [{sheetName}]";

                    var dataTable = new DataTable();
                    dataTable.TableName = sheetName;

                    OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command);
                    dataAdapter.Fill(dataTable);

                    dataSet.Tables.Add(dataTable);
                }
            }
            return dataSet;
        }
    }
}
