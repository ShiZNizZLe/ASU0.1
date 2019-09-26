using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.XlsIO;
using System.IO;
using System.Reflection;

namespace Excel_to_database_3
{
    public class Excel
    {
        //Creates a new instance for ExcelEngine
        private ExcelEngine excelEngine;
        private IWorkbook workbook;

        public Excel()
        {
            this.excelEngine = new ExcelEngine();

            //Initialize Application
            IApplication application = excelEngine.Excel;

            //Set default application version as Excel 2016
            application.DefaultVersion = ExcelVersion.Excel2016;
            this.workbook = this.excelEngine.Excel.Workbooks.Open("C:\\Users\\Koen\\Downloads\\test_sheet_asu.xlsm", ExcelParseOptions.Default, false, "ASU");
        }

        
        public void DisplayText()
        {
            
            IWorksheet worksheet = this.workbook.Worksheets[2];
            IMigrantRange migrantRange = worksheet.MigrantRange;
            worksheet.UsedRangeIncludesFormatting = false;
            int rowCount = worksheet.UsedRange.LastRow;
            int colCount = worksheet.UsedRange.LastColumn;
            List<string> rowsList = new List<string>();
            List<string> rowsContentList = new List<string>();
            for (int i = 8; i <= rowCount; i++)
            {
                //string rowValue = migrantRange.DisplayText.ToString();
                //rowsList.Append<string>(rowValue);
                string value = "";
                for (int j = 1; j <= colCount; j++)
                {

                    migrantRange.ResetRowColumn(i, j);
                    value = value + " - " +  migrantRange.DisplayText;                    

                }

                Console.WriteLine(value);
            }
            //IRange[] cells = worksheet["A8:C40"].Cells;


            //give value of each cell in range
            /*foreach (IRange cell in cells)
            {
                string cellValue = cell.DisplayText.ToString();
                
                //Console.WriteLine(cellValue);
            }*/

            //give value of each row in range
            //IRange[] rows = worksheet["A8:"].Rows;
            
     
            /*foreach (IRange row in rows)
            {
                row.ToArray();
                display value of the first cell of the row
                string rowValue = row.DisplayText.ToString();
                Console.WriteLine(rowValue);

            
            }*/

            //string displayText = worksheet.Range["C8:C10"].DisplayText;
            //getrange of cells and display a part of them
            //IRange range = worksheet.Range[1, 8, 16, 534];
            //string displayText = range[1, 9, 4, 20].Text;
            //string displayText = worksheet.GetValueRowCol(8, 3).ToString();

            worksheet.UsedRangeIncludesFormatting = false;
            Console.WriteLine(worksheet.UsedRange.AddressLocal); 
            Close();
        }

        private void FindHeaders(IRange row, int rowLength)
        {
            for(int i=0;  i < rowLength; i++)
            {
                migrantRange.ResetRowColumn(, j);
                String value =
                switch
            }
        }

        public void Close()
        {
            this.workbook.Close(false);
        }

    }
}
