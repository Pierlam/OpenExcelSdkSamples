using OpenExcelSdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdkSamplesApp;

public class GetCellValueRowByRow
{
    public static void Run()
    {
        ExcelProcessor proc = new ExcelProcessor();

        Console.WriteLine("-> GetCellValueRowByRow.Run:");

        // open an excel file
        string filename = @".\ExcelFiles\GetCellValueRowByRow.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        // get the first sheet of the excel file
        ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

        int lastRowIdx= proc.GetLastRowIndex(excelSheet);
        Console.WriteLine($"LastRowIndex: {lastRowIdx}");

        //--scan each row
        for (int i=1; i< lastRowIdx;i++) 
        {
            Console.WriteLine($"---");

            if (i == 9)
            {

            }
            // get the row by index, if the row doesn't exists, row is null, it's not an error
            ExcelRow excelRow = proc.GetRowAt(excelSheet, i);
            if (excelRow == null)
            {
                Console.WriteLine($"Row {i}, is empty, has no cell.");
                continue;
            }


            //--scan each cell in the row
            foreach (ExcelCell excelCell in proc.GetRowCells(excelSheet,excelRow))
            {
                // get the cell value, if the cell doesn't exists, cell is null, it's not an error
                if(excelCell == null) continue;

                ExcelCellValue excelCellValue = proc.GetCellValue(excelSheet, excelCell);

                if (excelCellValue.IsEmpty)
                {
                    Console.WriteLine($"Row {i}, Cell {excelCell.Cell.CellReference} is empty/blank.");
                    continue;
                }

                if (excelCellValue.CellType== ExcelCellType.String)
                {
                    Console.WriteLine($"Row {i}, Cell {excelCell.Cell.CellReference}, Type: String, Value:{excelCellValue.StringValue} ");
                    continue;
                }

                if (excelCellValue.CellType == ExcelCellType.Integer)
                {
                    Console.WriteLine($"Row {i}, Cell {excelCell.Cell.CellReference}, Type: Integer, Value:{excelCellValue.IntegerValue} ");
                    continue;
                }

                if (excelCellValue.CellType == ExcelCellType.Double)
                {
                    // can be a currency
                    if(excelCellValue.Currency==null)
                       Console.WriteLine($"Row {i}, Cell {excelCell.Cell.CellReference}, Type: Double, Value:{excelCellValue.DoubleValue} ");
                    else
                        Console.WriteLine($"Row {i}, Cell {excelCell.Cell.CellReference}, Type: Currency, Value:{excelCellValue.DoubleValue}, CurrencyName: {excelCellValue.Currency.Name}");

                    continue;
                }

                if (excelCellValue.CellType == ExcelCellType.DateOnly)
                {
                    Console.WriteLine($"Row {i}, Cell {excelCell.Cell.CellReference}, Type: DateOnly, Value:{excelCellValue.DateOnlyValue} ");
                    continue;
                }

                if (excelCellValue.CellType == ExcelCellType.DateTime)
                {
                    Console.WriteLine($"Row {i}, Cell {excelCell.Cell.CellReference}, Type: DateTime, Value:{excelCellValue.DateTimeValue} ");
                    continue;
                }

                if (excelCellValue.CellType == ExcelCellType.TimeOnly)
                {
                    Console.WriteLine($"Row {i}, Cell {excelCell.Cell.CellReference}, Type: TimeOnly, Value:{excelCellValue.TimeOnlyValue} ");
                    continue;
                }

                // TODO:  other types
            }
        }


        proc.CloseExcelFile(excelFile);

        /* Display result in console:
         * TODO:
        */

    }

}
