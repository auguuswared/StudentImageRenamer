using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        string excelPath = @"D:\SMCTUT_RNIMG\FORMAT_EXCEL.xlsx";
        //string sourceImageFolder = @"D:\Augusware\SMCTUT_RNIMG\Student Image\";
        string outputImageRoot = @"D:\SMCTUT_RNIMG\RenamedImages1\";
        string baseImagePath = @"D:\SMCTUT_RNIMG\Student Image\";
        if (!File.Exists(excelPath))
        {
            Console.WriteLine("Excel file not found.");
            return;
        }

        using (var workbook = new XLWorkbook(excelPath))
        {
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

            foreach (var row in rows)
            {
                string name = row.Cell(2).GetValue<string>().Trim();
                string registerNo = row.Cell(3).GetValue<string>().Trim();
                string photoUrl = row.Cell(4).GetValue<string>().Trim();
                string className = row.Cell(5).GetValue<string>().Trim();

                if (string.IsNullOrWhiteSpace(registerNo) || string.IsNullOrWhiteSpace(photoUrl) || string.IsNullOrWhiteSpace(className))
                    continue;

                // Extract folder (e.g., "Botany") and file name (e.g., "URB020.jpg")
                //Uri uri = new Uri(photoUrl);
                //string[] segments = uri.Segments;
                //string folderName = segments.Length >= 6 ? segments[5].TrimEnd('/') : ""; // e.g., "Botany"
                //string photoFileName = segments.Last(); // e.g., "URB020.jpg"

                // Extract folder (e.g., "Botany") and file name (e.g., "URB020.jpg")
                photoUrl = photoUrl.Replace("\\", "/"); // Fix the URI format

                Uri uri = new Uri(photoUrl);
                string[] segments = uri.Segments;

                string folderName = segments.Length >= 6 ? segments[5].TrimEnd('/') : ""; // e.g., "Botany"
                string photoFileName = segments.Last(); // e.g., "URBO20.jpg"

                if (string.IsNullOrWhiteSpace(folderName) || string.IsNullOrWhiteSpace(photoFileName))
                {
                    Console.WriteLine($"Invalid photo URL format: {photoUrl}");
                    continue;
                }

                string sourceImagePath = Path.Combine(baseImagePath, folderName, photoFileName);
                if (!File.Exists(sourceImagePath))
                {
                    Console.WriteLine($"Image not found for: {photoFileName} at {sourceImagePath}");
                    continue;
                }

                // Prepare output
                string classFolderPath = Path.Combine(outputImageRoot, className);
                Directory.CreateDirectory(classFolderPath);

                string newFilePath = Path.Combine(classFolderPath, registerNo + Path.GetExtension(photoFileName));

                if (!File.Exists(newFilePath))
                {
                    File.Copy(sourceImagePath, newFilePath);
                    Console.WriteLine($"Copied: {photoFileName} → {className}\\{registerNo + Path.GetExtension(photoFileName)}");
                }
                else
                {
                    Console.WriteLine($"Skipped: {registerNo + Path.GetExtension(photoFileName)} already exists.");
                }
            }
        }

        Console.WriteLine("Process Completed.");
    }
}
