using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        // Set the license context for EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        // Specify the path to your Excel file
        string excelFilePath = "/Users/apple/Documents/Book2.xlsx";// choose your file

        // Read RequestIDs and SupplierIDs from each sheet into HashSet
        HashSet<(string, string)> idsTest1 = ReadIds(excelFilePath, "test1", "RequestID", "SupplierID");
        HashSet<(string, string)> idsTest2 = ReadIds(excelFilePath, "test2", "RequestID", "SupplierID");

        // Find IDs in test2 but not in test1
        var idsNotInTest1 = idsTest2.Except(idsTest1).ToList();

        // Apply conditional formatting to highlight differences in the "test2" sheet
        ApplyConditionalFormatting(excelFilePath, "test2", "RequestID", "SupplierID", idsNotInTest1);

        Console.WriteLine("Differences highlighted successfully.");
    }

    static HashSet<(string, string)> ReadIds(string filePath, string sheetName, string requestIDColumnName, string supplierIDColumnName)
    {
        var ids = new HashSet<(string, string)>();

        try
        {
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = excelPackage.Workbook.Worksheets[sheetName];

                // Assuming the first row contains headers
                int requestIDColumnIndex = GetColumnIndex(worksheet, requestIDColumnName);
                int supplierIDColumnIndex = GetColumnIndex(worksheet, supplierIDColumnName);

                if (requestIDColumnIndex != -1 && supplierIDColumnIndex != -1)
                {
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var requestID = worksheet.Cells[row, requestIDColumnIndex].Text.Trim();
                        var supplierID = worksheet.Cells[row, supplierIDColumnIndex].Text.Trim();
                        ids.Add((requestID, supplierID));
                    }
                }
                else
                {
                    Console.WriteLine($"Column(s) not found in sheet {sheetName}.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error reading sheet {sheetName} in file {filePath}: {ex.Message}");
        }

        return ids;
    }

    static void ApplyConditionalFormatting(string filePath, string sheetName, string requestIDColumnName, string supplierIDColumnName, List<(string, string)> differingIds)
    {
        try
        {
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = excelPackage.Workbook.Worksheets[sheetName];

                // Assuming the first row contains headers
                int requestIDColumnIndex = GetColumnIndex(worksheet, requestIDColumnName);
                int supplierIDColumnIndex = GetColumnIndex(worksheet, supplierIDColumnName);

                if (requestIDColumnIndex != -1 && supplierIDColumnIndex != -1)
                {
                    int rowCount = worksheet.Dimension.Rows;

                    foreach (var differingId in differingIds)
                    {
                        var (requestID, supplierID) = differingId;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            var cellRequestID = worksheet.Cells[row, requestIDColumnIndex].Text.Trim();
                            var cellSupplierID = worksheet.Cells[row, supplierIDColumnIndex].Text.Trim();

                            if (cellRequestID == requestID && cellSupplierID == supplierID)
                            {
                                // Apply conditional formatting to highlight differing cells
                                worksheet.Cells[row, requestIDColumnIndex].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                worksheet.Cells[row, requestIDColumnIndex].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                                worksheet.Cells[row, supplierIDColumnIndex].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                worksheet.Cells[row, supplierIDColumnIndex].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);

                                break; // No need to continue searching in the same row
                            }
                        }
                    }

                    // Save the Excel file
                    excelPackage.Save();
                }
                else
                {
                    Console.WriteLine($"Column(s) not found in sheet {sheetName}.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error applying conditional formatting in sheet {sheetName} in file {filePath}: {ex.Message}");
        }
    }

    static int GetColumnIndex(ExcelWorksheet worksheet, string columnName)
    {
        // Assuming the first row contains headers
        int headerRow = 1;

        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
        {
            if (worksheet.Cells[headerRow, col].Text.Trim().Equals(columnName, StringComparison.OrdinalIgnoreCase))
            {
                return col;
            }
        }

        return -1;
    }
}
