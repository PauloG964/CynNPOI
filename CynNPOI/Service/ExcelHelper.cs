using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

public class ExcelHelper
{
    public byte[] WriteExcel(string format)
    {
        IWorkbook workbook;

        if (format == ".xls")
            workbook = new HSSFWorkbook();
        else if (format == ".xlsx")
            workbook = new XSSFWorkbook();
        else
            throw new NotSupportedException("File extension not supported.");

        ISheet sheet = workbook.CreateSheet("MySheet");
        IRow header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Name");
        header.CreateCell(1).SetCellValue("Email");

        IRow data = sheet.CreateRow(1);
        data.CreateCell(0).SetCellValue("Paul Ojieh");
        data.CreateCell(1).SetCellValue("paul@example.com");

        using var ms = new MemoryStream();
        workbook.Write(ms);
        return ms.ToArray();
    }

    public List<Dictionary<string, string>> ReadExcel(Stream stream, string extension)
    {
        IWorkbook workbook;

        if (extension == ".xls")
            workbook = new HSSFWorkbook(stream); // BIFF format (Excel 97-2003)
        else if (extension == ".xlsx" || extension == ".xlsm")
            workbook = new XSSFWorkbook(stream); // OpenXML format (macros not preserved)
        else
            throw new NotSupportedException("File extension not supported.");

        ISheet sheet = workbook.GetSheetAt(0);
        var result = new List<Dictionary<string, string>>();
        var headerRow = sheet.GetRow(0);

        for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            IRow row = sheet.GetRow(rowIndex);
            if (row == null) continue;

            var rowData = new Dictionary<string, string>();
            for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
            {
                string header = headerRow.GetCell(colIndex)?.ToString() ?? $"Column{colIndex}";
                string value = row.GetCell(colIndex)?.ToString();
                rowData[header] = value;
            }
            result.Add(rowData);
        }

        return result;
    }
}
