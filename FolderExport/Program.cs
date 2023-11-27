using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlDirectoryGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            // Specify the path to the directory
            string directoryPath = @"D:\Books";

            //  XLWorkbook wb = new XLWorkbook();


            // Create a new Excel workbook
            //using (SpreadsheetDocument spreadsheetDocument =
            //       SpreadsheetDocument.Create("DirectoryList.xlsx", SpreadsheetDocumentType.Workbook))
            //{
            //    // Create a new worksheet
            //    WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            //    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            //    Worksheet worksheet = worksheetPart.Worksheet;

            //    // Create a new sheetData element
            //    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            //    if (sheetData == null)
            //    {
            //        sheetData = worksheet.AppendChild(new SheetData());
            //    }

            //    // Write the header row
            //    Row row = sheetData.AppendChild(new Row());
            //    Cell folderNameCell = row.AppendChild(new Cell());
            //    folderNameCell.CellValue = new CellValue("Folder Name");

            //    Cell fileNameCell = row.AppendChild(new Cell());
            //    fileNameCell.CellValue = new CellValue("File Name");

            // Get all folders, subfolders, and files in the specified path
            List<string> folders = new List<string>();
            GetFolders(directoryPath, folders);

            // Write the folders and files to the Excel worksheet
            int rowNumber = 2;
            foreach (string folder in folders)
            {
                //row = sheetData.AppendChild(new Row());

                //  Cell folderCell = row.AppendChild(new Cell());
                //folderCell.CellValue = new CellValue(folder);
                Console.WriteLine(folder);

                string[] files = Directory.GetFiles(folder);
                foreach (string file in files)
                {
                    //  row = sheetData.AppendChild(new Row());

                    //  Cell cell = row.AppendChild(new Cell());
                    // cell.CellValue = new CellValue(file);
                    Console.WriteLine(file);
                    rowNumber++;
                }

                rowNumber++;
                //  }
            }
        }

        static void GetFolders(string directoryPath, List<string> folders)
        {
            folders.Add(directoryPath);

            string[] subDirectories = Directory.GetDirectories(directoryPath);
            foreach (string subDirectory in subDirectories)
            {
                GetFolders(subDirectory, folders);
            }
        }
    }
}
