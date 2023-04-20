using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml;

namespace ExcelUnlocker
{
    class Program
    {
        static void Main(string[] args)
        {
            printBanner();
            string? inputFile = null;


            Console.WriteLine("Input file name (es Example.xlsx): ");
            do
            {
                inputFile = Console.ReadLine();

            } while (String.IsNullOrWhiteSpace(inputFile));





            string outputFile = "unlocked_file.xlsx";
            string tempZipFile = "temp.zip";
            string tempZipFile2 = "temp2.zip";

            if (File.Exists(inputFile))
            {
                // Rename file to .zip
                File.Move(inputFile, tempZipFile);

                // Extract the zip archive
                ZipFile.ExtractToDirectory(tempZipFile, "temp");

                // Go through all xml files in folder "xl/worksheets"
                string[] xmlFiles = Directory.GetFiles("temp/xl/worksheets", "*.xml", SearchOption.AllDirectories);
                foreach (string xmlFile in xmlFiles)
                {
                    // Read the xml file content
                    string xmlContent = File.ReadAllText(xmlFile);
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(xmlContent);

                    doc.GetElementsByTagName("sheetProtection")[0]?.RemoveAll();
                    doc.Save(xmlFile);

                    // Remove the entire xml row starting with "<sheetProtection"
                    //xmlContent = Regex.Replace(xmlContent, "<sheetProtection.* />", "", RegexOptions.Singleline);

                    //// Save the modified xml file
                    //File.WriteAllText(xmlFile, xmlContent);
                }

                // Open the "workbook.xml" file in the xl folder
                string workbookXmlFile = "temp/xl/workbook.xml";
                string workbookXmlContent = File.ReadAllText(workbookXmlFile);

                // Change the "lockStructure" parameter from 1 to 0 in the "workbookProtection" row
                workbookXmlContent = Regex.Replace(workbookXmlContent, "lockStructure=\"1\"", "lockStructure=\"0\"");

                // Save the modified workbook.xml file
                File.WriteAllText(workbookXmlFile, workbookXmlContent);

                // Create the modified zip archive
                ZipFile.CreateFromDirectory("temp", tempZipFile2);

                // Rename the zip archive to the desired output file
                File.Move(tempZipFile2, outputFile);

                // Delete the extracted files
                Directory.Delete("temp", true);
                File.Delete(tempZipFile);

                Console.WriteLine("Excel file unlocked successfully. Press any key to exit");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine("Input file not found.");
                Console.ReadLine();

            }
        }
        public static void printBanner()
        {
            Random rnd = new Random();
            int num = rnd.Next(16);
            Console.ForegroundColor = (ConsoleColor)num;
            Console.WriteLine("--------------------------------------------------------------");
            Console.WriteLine(@" _____              _ _   _       _            _");
            Console.WriteLine(@"| ____|_  _____ ___| | | | |_ __ | | ___   ___| | _____ _ __");
            Console.WriteLine(@"|  _| \ \/ / __/ _ \ | | | | '_ \| |/ _ \ / __| |/ / _ \ '__|");
            Console.WriteLine(@"| |___ >  < (_|  __/ | |_| | | | | | (_) | (__|   <  __/ |");
            Console.WriteLine(@"|_____/_/\_\___\___|_|\___/|_| |_|_|\___/ \___|_|\_\___|_|");
            Console.WriteLine("--------------------------------------------------------------");
            Console.WriteLine("Made by Simo, all right reserved ©" + DateTime.Now.Year);
            Console.WriteLine("\n");
            Console.ResetColor();

        }
    }


}
