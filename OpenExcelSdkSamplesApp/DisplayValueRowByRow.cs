using OpenExcelSdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdkSamplesApp;

public class DisplayValueRowByRow
{
    public static void Run()
    {
        ExcelProcessor proc = new ExcelProcessor();

        // open an excel file
        string filename = @".\ExcelFiles\BasicSample.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        // get the first sheet of the excel file
        ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

        int lastRowIdx= proc.GetLastRowIndex(excelSheet);

        // scan each row
        for (int i=0; i< lastRowIdx;i++) 
        {
            // get the row by index, if the row doesn't exists, row is null, it's not an error
            ExcelRow excelRow = proc.GetRowAt(excelSheet, i);
            if(excelRow == null) 
                 continue; 

            // scan each cell in the row
            foreach (ExcelCell excelCell in proc.GetRowCells(excelSheet,excelRow))
            {
                // get the cell value, if the cell doesn't exists, cell is null, it's not an error 
                // TODO:
                //object cellValue = proc.GetCellValue(excelCell);

                //Console.WriteLine($"Row {i}, Cell {excelCell.CellIndex}, Value: {cellValue}, Type: {cellValue?.GetType().Name}");
            }
        }

        //==A2: hello - string
        // get a cell, if the cell doesn't exists, cell is null, it's not an error 
        //ExcelCell excelCell = proc.GetCellAt(excelSheet, "A2");

    }

}
