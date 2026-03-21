using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcelSdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdkSamplesApp;

/// <summary>
/// Read some  cell values from an excel file, and display the type and value of the cell.
/// </summary>
public class GetCellValue
{
    public static void GetBasicValue()
    {
        ExcelProcessor proc = new ExcelProcessor();

        Console.WriteLine("-> GetCellValue.GetBasicValue:");

        // open an excel file
        string filename = @".\ExcelFiles\GetCellValue.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        // get the first sheet of the excel file
        ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

        //==A2: hello - string
        // get a cell, if the cell doesn't exists, cell is null, it's not an error 
        ExcelCell excelCell = proc.GetCellAt(excelSheet, "A2");

        // get the type and the value of cell
        ExcelCellValue  excelCellValue = proc.GetCellValue(excelSheet, excelCell);
        Console.WriteLine("A2: Value: {0}, cell type: {1}", excelCellValue.StringValue, excelCellValue.CellType.ToString());

        //==A3: 12 - Integer
        excelCell = proc.GetCellAt(excelSheet, "A3");
        excelCellValue = proc.GetCellValue(excelSheet, excelCell);
        Console.WriteLine("A3: Value: {0}, cell type: {1}", excelCellValue.IntegerValue, excelCellValue.CellType.ToString());

        //==A4: 34,45 - Double
        excelCell = proc.GetCellAt(excelSheet, "A4");
        excelCellValue = proc.GetCellValue(excelSheet, excelCell);
        Console.WriteLine("A4: Value: {0}, cell type: {1}", excelCellValue.DoubleValue, excelCellValue.CellType.ToString());

        //==A5: 12/10/2025 - DateOnly
        excelCell = proc.GetCellAt(excelSheet, "A5");
        excelCellValue = proc.GetCellValue(excelSheet, excelCell);
        Console.WriteLine("A5: Value: {0}, cell type: {1}", excelCellValue.DateOnlyValue.ToString(), excelCellValue.CellType.ToString());

        //==A6: 03:45 - TimeOnly
        excelCell = proc.GetCellAt(excelSheet, "A6");
        excelCellValue = proc.GetCellValue(excelSheet, excelCell);
        Console.WriteLine("A6: Value: {0}, cell type: {1}", excelCellValue.TimeOnlyValue.ToString(), excelCellValue.CellType.ToString());

        proc.CloseExcelFile(excelFile);


        /* Display result in console:
        A2: Value: hello, cell type: String
        A3: Value: 12, cell type: Integer
        A4: Value: 34,45, cell type: Double
        A5: Value: 12/10/2025, cell type: DateOnly
        A6: Value: 03:45, cell type: TimeOnly
        */
    }


    public static void GetCurrencyValue()
    {
        ExcelProcessor proc = new ExcelProcessor();

        Console.WriteLine("-> GetCellValue.GetCurrencyValue:");

        // open an excel file
        string filename = @".\ExcelFiles\GetCellValue.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        // get the first sheet of the excel file
        ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

        //==A7: 1 200 €	- currency, euro
        ExcelCell excelCell = proc.GetCellAt(excelSheet, "A7");
        ExcelCellValue excelCellValue = proc.GetCellValue(excelSheet, excelCell);
        Console.WriteLine("A7: Value: {0}, cell type: {1}, CurrencyName: {2}", excelCellValue.DoubleValue, excelCellValue.CellType.ToString(), excelCellValue.Currency.Name);

        //==A8:  $3 400,00 - currency,  US Dollar
        excelCell = proc.GetCellAt(excelSheet, "A8");
        excelCellValue = proc.GetCellValue(excelSheet, excelCell);
        Console.WriteLine("A8: Value: {0}, cell type: {1}, CurrencyName: {2}", excelCellValue.DoubleValue, excelCellValue.CellType.ToString(), excelCellValue.Currency.Name);

        proc.CloseExcelFile(excelFile);

        /* Display result in console:
        => OpenExcelSdkSamplesApp:
        A7: Value: 1200, cell type: Double, CurrencyName: Euro
        A8: Value: 3400, cell type: Double, CurrencyName: UsDollar        
        */

    }
}
