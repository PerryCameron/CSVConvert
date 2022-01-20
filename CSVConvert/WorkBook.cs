using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;

namespace CSVConvert
{
    class WorkBook : XSSFWorkbook
    {
        List<Product> productErrorList;
        XSSFCellStyle borderedCellStyle;
        public string OutputDir { get; set; }
        public WorkBook(List<Product> productErrorList)
        {
            this.productErrorList = productErrorList;
        }

        public void CreateXSSFErrorFile() { 
        // set a font and size for excel, I have it at default
        XSSFFont myFont = (XSSFFont)this.CreateFont();
        myFont.FontHeightInPoints = 11;
        myFont.FontName = "Calibri";
        // creates a style, set to default
        this.borderedCellStyle = (XSSFCellStyle)this.CreateCellStyle();
        borderedCellStyle.SetFont(myFont);
        borderedCellStyle.BorderLeft = BorderStyle.None;
        borderedCellStyle.BorderTop = BorderStyle.None;
        borderedCellStyle.BorderRight = BorderStyle.None;
        borderedCellStyle.BorderBottom = BorderStyle.None;
        borderedCellStyle.VerticalAlignment = VerticalAlignment.Center;
        ISheet Sheet = this.CreateSheet("Error");
        //create header row
        CreateRow(Sheet.CreateRow(0), productErrorList[0]);
        // we no longer need this row
        productErrorList.RemoveAt(0);
        // add content
        PopulateRows(Sheet);
            using (var fileData = new FileStream(OutputDir + "Error.xlsx", FileMode.Create))
            {
                this.Write(fileData);
            }

        }

        // creates a row for every row in error list
        private void PopulateRows(ISheet Sheet)
        {
            for (int i = 1; i < productErrorList.Count; i++)
            {
                CreateRow(Sheet.CreateRow(i), productErrorList[i - 1]);
            }

        }

        // populates a row of cells at one time
        private void CreateRow(IRow CurrentRow, Product p)
        {
            CreateCell(CurrentRow, 0, p.pid, borderedCellStyle);
            CreateCell(CurrentRow, 1, p.productId, borderedCellStyle);
            CreateCell(CurrentRow, 2, p.manufacturerName, borderedCellStyle);
            CreateCell(CurrentRow, 3, p.manufacturerPN, borderedCellStyle);
            CreateCell(CurrentRow, 4, p.cost, borderedCellStyle);
            CreateCell(CurrentRow, 5, p.coo, borderedCellStyle);
            CreateCell(CurrentRow, 6, p.description, borderedCellStyle);
            CreateCell(CurrentRow, 7, p.upc, borderedCellStyle);
            CreateCell(CurrentRow, 8, p.uom, borderedCellStyle);
        }

        // creates a cell with properties and value
        private void CreateCell(IRow CurrentRow, int CellIndex, string Value, XSSFCellStyle Style)
        {
            ICell Cell = CurrentRow.CreateCell(CellIndex);
            Cell.SetCellValue(Value);
            Cell.CellStyle = Style;
        }
    }
}
