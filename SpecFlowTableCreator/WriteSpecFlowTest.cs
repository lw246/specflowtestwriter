using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpecFlowTableCreator
{
    public class WriteSpecFlowTest
    {
        public string SpecFlowTest { get { return _specflowTest.ToString(); } }

        private StringBuilder _specflowTest;
        private string _filePath;
        private string _collectionName;

        public WriteSpecFlowTest()
        {
            _specflowTest = new StringBuilder();
        }

        public static WriteSpecFlowTest New()
        {
            return new WriteSpecFlowTest();
        }

        public WriteSpecFlowTest UsingFile(string filePath)
        {
            _filePath = filePath;
            return this;
        }

        public WriteSpecFlowTest ForCollection(string collectionName)
        {
            _collectionName = collectionName;
            return this;
        }

        public WriteSpecFlowTest CreateTestGivenStatements()
        {
            _specflowTest.Append($"Given the ispac file {_collectionName}.ispac");
            _specflowTest.AppendLine($"And the collection {_collectionName} is deployed");
            return this;
        }

        public WriteSpecFlowTest CreateTestWhenStatement()
        {
            _specflowTest.AppendLine($"When the {_collectionName} is executed");
            return this;
        }

        public WriteSpecFlowTest CreateThenStatements()
        {
            Dictionary<string, string> specflowTables = SpecFlowTableWriter.New()
                .ReadDataFromFile(_filePath)
                .Create();

            foreach (var specflowTable in specflowTables)
            {
                WriteSpecFlowTableThen(specflowTable);
            }

            return this;
        }

        private void WriteSpecFlowTableThen(KeyValuePair<string, string> specflowTable)
        {
            _specflowTest.AppendLine($"Then the Excel file '{Path.GetFileName(_filePath)}' contains sheet '{specflowTable.Key}' with contents");
            _specflowTest.Append(specflowTable.Value + Environment.NewLine);
        }

        public void SaveToFile(string savePath)
        {
            if (_specflowTest is null)
            {
                throw new FieldAccessException("No SpecFlow test has been created yet");
            }
            if (!Directory.Exists(savePath))
            {
                throw new DirectoryNotFoundException($"Directory {savePath} doesn't exist, please create it first");
            }

            File.WriteAllText(Path.Combine(savePath, $"{_collectionName}.feature"), _specflowTest.ToString());
        }
    }
}
