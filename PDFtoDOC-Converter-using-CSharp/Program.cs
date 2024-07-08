using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Xml.Linq;
using Aspose.Pdf;

namespace PDFtoDOC_Converter_using_CSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourceFolder = ConfigurationManager.AppSettings["SourceFolder"];
            string destinationFolder = ConfigurationManager.AppSettings["DestinationFolder"];
            string connectionString = ConfigurationManager.ConnectionStrings["FileMoveHistoryDB"].ConnectionString;

            // Check if source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder {sourceFolder} does not exist.");
                return;
            }

            // Ensure destination folder exists
            if (!Directory.Exists(destinationFolder))
            {
                try
                {
                    Directory.CreateDirectory(destinationFolder);
                    Console.WriteLine($"Created folder: {destinationFolder}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error creating folder {destinationFolder}: {ex.Message}");
                    return;
                }
            }

            // Get all PDF files from the source folder
            string[] pdfFiles = Directory.GetFiles(sourceFolder, "*.pdf");
            List<FileMoveRecord> moveRecords = new List<FileMoveRecord>();

            foreach (string file in pdfFiles)
            {
                try
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    string destFile = Path.Combine(destinationFolder, $"{fileName}.doc");

                    // Convert PDF to DOC
                    if (ConvertPdfToDoc(file, destFile))
                    {
                        File.Delete(file);
                        Console.WriteLine($"Converted and moved {file} to {destFile}");
                        moveRecords.Add(new FileMoveRecord(fileName, "DOC", sourceFolder, destFile));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error moving file {file}: {ex.Message}");
                }
            }

            // Insert move history into SQL table
            DAL.InsertMoveHistory(moveRecords, connectionString);

            // Print summary
            Console.WriteLine("\nMove Process Summary:");
            Console.WriteLine($"Total files converted and moved: {moveRecords.Count}");
        }

        static bool ConvertPdfToDoc(string sourceFile, string destinationFile)
        {
            try
            {
                Document pdfDocument = new Document(sourceFile);
                pdfDocument.Save(destinationFile, Aspose.Pdf.SaveFormat.Doc);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting file {sourceFile} to DOC: {ex.Message}");
                return false;
            }
        }
    }

    public class FileMoveRecord
    {
        public string FileName { get; }
        public string FileType { get; }
        public string SourcePath { get; }
        public string DestinationPath { get; }

        public FileMoveRecord(string fileName, string fileType, string sourcePath, string destinationPath)
        {
            FileName = fileName;
            FileType = fileType;
            SourcePath = sourcePath;
            DestinationPath = destinationPath;
        }
    }
}
