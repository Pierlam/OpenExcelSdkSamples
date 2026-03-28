using OpenExcelSdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcelSdkSamplesApp;

public class ScanDataRowCells
{
    /// <summary>
    /// Scans an Excel file to retrieve and display the addresses of non-empty cells in existing rows.
    /// </summary>
    /// <remarks>This method opens an Excel file, retrieves the first sheet, and iterates through each
    /// existing row to print the addresses of cells that contain values. It does not display empty rows or
    /// cells.</remarks>
    public static void ScanOnlyExistingRowsAndExistingCells()
    {
        ExcelProcessor proc = new ExcelProcessor();

        Console.WriteLine("Scan only existing rows and cells");

        // open an excel file
        string filename = @"ExcelFiles\scanDatatable.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        // get the first sheet of the excel file
        ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

        int lastRowIdx = proc.GetLastRowIndex(excelSheet);
        Console.WriteLine($"LastRowIndex: {lastRowIdx}");


        //--scan each existing row
        for (int r = 1; r <= lastRowIdx; r++)
        {
            Console.WriteLine("---");
            Console.WriteLine($"Row idx:{r}");

            // get the row by index, if the row doesn't exists, row is null, it's not an error
            ExcelRow excelRow = proc.GetRowAtIndex(excelSheet, r);

            // get cells of the row
            List<ExcelCell> listCells = proc.GetRowCells(excelSheet, excelRow);

            // scan each cell of the row
            foreach (ExcelCell cell in listCells)
            {
                Console.WriteLine($"Cell addr: {cell.Cell.CellReference} has a value");
            }
        }

        proc.CloseExcelFile(excelFile);

        /*
    => OpenExcelSdk DevApp:
    ScanDataTableWayOne: Does NOT display empty rows and cells!
    LastRowIndex: 4
    ---
    Row idx:1
    Cell addr: A1 has a value
    Cell addr: B1 has a value
    Cell addr: C1 has a value
    ---
    Row idx:2
    Cell addr: A2 has a value
    Cell addr: B2 has a value
    Cell addr: C2 has a value
    ---
    Row idx:3
    Cell addr: A4 has a value
    Cell addr: C4 has a value
    ---
    Row idx:4
    Cell addr: A6 has a value
    Cell addr: B6 has a value
    => Ok, Ends.
        */
    }

    /// <summary>
    /// Scans an Excel data table and writes the addresses of all non-empty cells in each row to the console, indicating
    /// empty rows without displaying cell content.
    /// </summary>
    /// <remarks>This method opens a specified Excel file, retrieves the first worksheet, and iterates through
    /// its rows. For each row, it outputs the row address and the addresses of any non-empty cells. If a row contains
    /// no cells, only the row address is displayed. The method is intended for diagnostic or inspection purposes and
    /// requires that the Excel file exists at the given path.</remarks>
    public static void ScanAllRowsOnlyExistingCells()
    {

        ExcelProcessor proc = new ExcelProcessor();

        Console.WriteLine("Scan all rows (with empty ones) but only existing cells");

        // open an excel file
        string filename = @"ExcelFiles\scanDatatable.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        // get the first sheet of the excel file
        ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

        int lastRowAddr = proc.GetLastRowAddress(excelSheet);
        Console.WriteLine($"LastRowAddress: {lastRowAddr}");


        //--scan each existing row
        for (int r = 1; r <= lastRowAddr; r++)
        {
            Console.WriteLine("---");
            Console.WriteLine($"Row addr:{r}");

            // get cells of the row
            List<ExcelCell> listCells = proc.GetRowCellsAtAddress(excelSheet, r);

            // scan each cell of the row
            foreach (ExcelCell cell in listCells)
            {
                Console.WriteLine($"Cell addr: {cell.Cell.CellReference}");
            }
        }

        proc.CloseExcelFile(excelFile);

        /*
         * ScanDataTableWayTwo: Display empty rows but not empty cells
    LastRowAddress: 6
    ---
    Row addr:1
    Cell addr: A1
    Cell addr: B1
    Cell addr: C1
    ---
    Row addr:2
    Cell addr: A2
    Cell addr: B2
    Cell addr: C2
    ---
    Row addr:3
    ---
    Row addr:4
    Cell addr: A4
    Cell addr: C4
    ---
    Row addr:5
    ---
    Row addr:6
    Cell addr: A6
    Cell addr: B6
     */
    }


    /// <summary>
    /// Scans an Excel file and displays the contents of each row and cell, highlighting empty rows and cells. 
    /// </summary>
    /// <remarks>This method opens a specified Excel file, iterates through its rows and columns, and outputs
    /// the status of each cell to the console, indicating whether it is empty or contains a value. Ensure that the file
    /// path is valid and accessible before calling this method.</remarks>
    public static void ScanAllRowsAndCells()
    {

        ExcelProcessor proc = new ExcelProcessor();

        Console.WriteLine("Scan all rows (with empty ones) and all cells (with null ones)");

        // open an excel file
        string filename = @"ExcelFiles\scanDatatable.xlsx";
        ExcelFile excelFile = proc.OpenExcelFile(filename);

        // get the first sheet of the excel file
        ExcelSheet excelSheet = proc.GetFirstSheet(excelFile);

        int lastRowAddr = proc.GetLastRowAddress(excelSheet);
        Console.WriteLine($"LastRowAddress: {lastRowAddr}");


        //--scan each existing row
        for (int r = 1; r <= lastRowAddr; r++)
        {
            Console.WriteLine("---");
            Console.WriteLine($"Row addr:{r}");

            int lastColAddr = proc.GetLastColAddress(excelSheet, r);

            for (int c = 1; c <= lastColAddr; c++)
            {
                ExcelCell cell = proc.GetCellAt(excelSheet, c, r);
                if (cell == null)
                {
                    Console.WriteLine($"Cell addr: Col:{c}, Row{r}: cell is empty");
                }
                else
                {
                    Console.WriteLine($"Cell addr: {cell.Cell.CellReference}: Cell has a value");
                }
            }
        }

        proc.CloseExcelFile(excelFile);

        /*
    ScanDataTableWayThree: Display empty rows and cells
    LastRowAddress: 6
    ---
    Row addr:1
    Cell addr: A1: Cell has a value
    Cell addr: B1: Cell has a value
    Cell addr: C1: Cell has a value
    ---
    Row addr:2
    Cell addr: A2: Cell has a value
    Cell addr: B2: Cell has a value
    Cell addr: C2: Cell has a value
    ---
    Row addr:3
    ---
    Row addr:4
    Cell addr: A4: Cell has a value
    Cell addr: Col:2, Row4: cell is empty
    Cell addr: C4: Cell has a value
    ---
    Row addr:5
    ---
    Row addr:6
    Cell addr: A6: Cell has a value
    Cell addr: B6: Cell has a value
    => Ok, Ends.      
         */
    }

}
