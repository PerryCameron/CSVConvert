using System;

namespace CSVConvert
{
    class Program
    {
        static void Main(string[] args)
        {
            // instantiate parser to extract and parse information from XLSX to List Arrays
            ProductParser parser = new ProductParser();
            // instantiate printer and pass Good List to write CSV files
            Printer printer = new Printer(parser.productList);
            // instantiate workbook to write bad rows to XLSX
            WorkBook workbook = new WorkBook(parser.productErrorList);
            // instantiate error checking
            Error error = new Error();

            // get our arguments
            string[] cleanargs = error.checkArgs(args);
            // make sure input file is correct
            error.FileIsCorrect(parser.InputPath = cleanargs[1]);
            // make sure directory exists
            if(error.directoryExists(printer.OutputDir = cleanargs[0]))
            workbook.OutputDir = printer.OutputDir;

            // stream data in, parse, extract and clean
            parser.TurnExcelIntoDTO();
            // make our CSV files
            printer.PrintCSVFiles();
            // make our XLSX file
            workbook.CreateXSSFErrorFile();
            // print the work we have done
            parser.PrintTotals();
        }
    }
}


