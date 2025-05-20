using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using WebApplication13.Context;

public class ExcelHelper
{
    public static void CreateExcelFromDataTable(DataTable dataTable, string filePath)
    {
        IWorkbook workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("SwipeData");

        // Create header row
        IRow headerRow = sheet.CreateRow(0);
        for (int i = 0; i < dataTable.Columns.Count; i++)
        {
            ICell cell = headerRow.CreateCell(i);
            cell.SetCellValue(dataTable.Columns[i].ColumnName);
        }

        // Create data rows
        for (int i = 0; i < dataTable.Rows.Count; i++)
        {
            IRow row = sheet.CreateRow(i + 1);
            for (int j = 0; j < dataTable.Columns.Count; j++)
            {
                ICell cell = row.CreateCell(j);
                cell.SetCellValue(dataTable.Rows[i][j].ToString());
            }
        }

        // Save the file
        using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
        {
            workbook.Write(fs);
        }
    }
}
