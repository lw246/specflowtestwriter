using System;
using System.IO;

namespace SpecFlowTableCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            AskForCollectionDetails(out string collectionName, out string filePath);

            WriteSpecFlowTest specflowTest = WriteSpecFlowTest.New()
                                .ForCollection(collectionName)
                                .UsingFile(filePath)
                                .CreateTestGivenStatements()
                                .CreateTestWhenStatement()
                                .CreateThenStatements();

            PrintTest(specflowTest);

            if (!PromptToSave())
            {
                return;
            }

            Save(specflowTest);

            return;
        }

        private static void PrintTest(WriteSpecFlowTest specflowTest)
        {
            Console.WriteLine("Do you wish to see the test?");

            if (Console.ReadLine().ToUpper().StartsWith("Y"))
            {
                Console.Write(specflowTest.SpecFlowTest);
            }
        }

        private static void Save(WriteSpecFlowTest specflowTest)
        {
            Console.WriteLine("Please enter the save location");
            var savePath = Path.GetFullPath(Console.ReadLine());

            while (!Directory.Exists(savePath))
            {
                Console.WriteLine($"Directory {savePath} doesn't exist. Please create it and press any key when done");
                savePath = Console.ReadLine();
            }

            specflowTest.SaveToFile(savePath);
        }

        private static bool PromptToSave()
        {
            Console.WriteLine("Would you like to save the test? Y/n");
            var response = Console.ReadLine().ToUpper();

            while (!response.StartsWith("N") && !response.StartsWith("Y"))
            {
                if (response.StartsWith("N"))
                {
                    return false;
                }

                if (!response.StartsWith("Y"))
                {
                    Console.WriteLine($"'{response}' is not a valid response.");
                    Console.WriteLine("Would you like to save the test? Y/n");
                    response = Console.ReadLine().ToUpper();
                }
            }

            return true;
        }

        private static void AskForCollectionDetails(out string collectionName, out string filePath)
        {
            Console.WriteLine("Enter collection name:");
            collectionName = Console.ReadLine();

            Console.WriteLine("Enter the path to the excel file");
            filePath = Path.GetFullPath(Console.ReadLine());

            while (!File.Exists(filePath))
            {
                Console.WriteLine($"No file found in location {filePath}. Please try again.");
                filePath = Path.GetFullPath(Console.ReadLine());
            }
        }
    }
}
