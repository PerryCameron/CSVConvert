using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace CSVConvert
{
    class ProductParser
    {
        // goodies
        public List<Product> productList { get; set; }
        // goodies that did not make the cut
        public List<Product> productErrorList { get; set; }
        // corrected records
        // input fullpath to file
        public String InputPath { get; set; }
        // output directory
        public String OutputDir { get; set; }
        // count of corrections
        public int Corrected { get; set; }
        // count of rows corrected
        public int rowsCorrected { get; set; }

        // open our workbook
        IWorkbook productBook = null;
        // there is only one sheet to worry about
        XSSFSheet sheet = null;

        public ProductParser()
        {
            this.productList = new List<Product>();
            this.productErrorList = new List<Product>();
        }

        public void TurnExcelIntoDTO()
        {
            this.productBook = new XSSFWorkbook(new FileStream(InputPath, FileMode.Open, FileAccess.Read));
            this.sheet = (XSSFSheet)productBook.GetSheetAt(0);

            // we will read from our xlsx and create our lists with one for loop
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                // create a dto
                Product product = AddRowToDTO(sheet.GetRow(rowIndex));
                // check for errors and put in appropriate list
                if (ProductHasErrors(product))
                    productErrorList.Add(product);
                else
                {
                    // with nice clean rows it is time to add the 20%
                    product.cost = (decimal.Multiply(Convert.ToDecimal(product.cost), Convert.ToDecimal(1.2))).ToString("#.##");
                    productList.Add(product);
                }
            }

        }

        private bool ProductHasErrors(Product product)
        {
            // make sure we have the correct data type, check the size, and place missing strings where applicable
            if (IntHasError(product.pid, 20)) return true;
            if (StringHasError(product.productId, 50)) return true;
            if (StringHasError(product.manufacturerName, 50)) return true;
            if (StringHasError(product.manufacturerPN, 50)) return true;
            if (DecimalHasError(product.cost, 20)) return true;
            if (StringHasError(product, 2, "TW")) return true;
            if (StringHasError(product.description, 300)) return true;
            if (StringHasError(product, 12, "")) return true;
            if (StringHasError(product, 2, "EA")) return true;
            return false;
        }

        private bool IntHasError(string pid, int size)
        {
            // nulls are considered a bad row for primary key
            if (pid == null) return true;
            // test to make sure this is an integer
            if (!int.TryParse(pid, out _)) return true;
            // make sure the length is within specification
            if (pid.Length > size) return true;
            return false;
        }

        private bool DecimalHasError(string price, int size)
        {
            // nulls are considered a bad row for cost
            if (price == null) return true;
            // check to make sure we have a proper number for currency
            if (!decimal.TryParse(price, out _)) return true;
            // make sure the length is within specification
            if (price.Length > size) return true;
            return false;
        }

        // buisness logic for handling coo, uom, upc columns
        private bool StringHasError(Product product, int size, string defaultAttribute)
        {
            if (defaultAttribute.Equals("TW"))
            {
                if (product.coo == null || product.coo.Equals(""))
                {
                    // add TW in if blank
                    product.coo = defaultAttribute;
                    Corrected++;
                }

                if (product.coo.Length > size) return true;

            }
            else if (defaultAttribute.Equals("EA"))
            {
                if (product.uom == null || product.upc.Equals(""))
                {
                    // add EA in if blank
                    product.uom = defaultAttribute;
                    Corrected++;
                }

                if (product.uom.Length > size) return true;
            }
            else
            {
                if (product.upc == null || product.upc.Equals(""))
                    product.upc = "";
                if (product.upc.Length > size) return true;
            }
            return false;
        }

        // don't usually overload, but this seemed as good a place as any
        private bool StringHasError(string field, int size)
        {
            if (field == null) return true;
            if (field.Length > size) return true;
            return false;
        }

        private Product AddRowToDTO(IRow row) // Ha! It rhymes.
        {
            return new Product(
                AddCellValue(row.GetCell(0)),
                AddCellValue(row.GetCell(1)),
                AddCellValue(row.GetCell(2)),
                AddCellValue(row.GetCell(3)),
                AddCellValue(row.GetCell(4)),
                AddCellValue(row.GetCell(5)),
                AddCellValue(row.GetCell(6)),
                AddCellValue(row.GetCell(7)),
                AddCellValue(row.GetCell(8)));
        }

        private string AddCellValue(ICell cell)
        {
            if (cell != null)
            {
                // removes any issues retrieving cell contents
                cell.SetCellType(CellType.String);
                return cell.StringCellValue;
            }
            // we found a null in the cell so lets return a null to the DTO
            else return null;
        }

        public void PrintTotals()
        {
            // we should not count header as bad row
            int BadRows = productErrorList.Count - 1;
            // total good records (after fixed)
            int GoodRows = productList.Count;
            // total rows
            int TotalRows = BadRows + GoodRows;
            Console.WriteLine("Satisfactory records: " + GoodRows);
            Console.WriteLine(BadRows + " record(s) could not be corrected");
            Console.WriteLine(Corrected + " errors corrected accross " + TotalRows + " rows");
        }
    }
}

