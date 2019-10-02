using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ShortageExport
{
    class ExcelFile
    {
        private string excelFilePath = string.Empty;
        private int rowNumber = 1; // define first row number to enter data in excel
        private int worksheetCount = 0;
        Excel.Application myExcelApplication;
        Excel.Workbook myExcelWorkbook;
        Excel.Worksheet ExcelShortageSheet;
        Excel.Worksheet ExcelAllItemsSheet;
        Excel.Worksheet ExcelShortTopsheet;
        public string ExcelFilePath
        {
            get { return excelFilePath; }
            set { excelFilePath = value; }
        }
        public int RowNumber {
            get { return rowNumber; }
            set { rowNumber = value; }
        }
        public Excel.Worksheet ShortageSheet
        {
            get { return ExcelShortageSheet; }
            set { ExcelShortageSheet = value; }
        }
        public Excel.Worksheet ShortageTopSheet
        {
            get { return ExcelShortTopsheet; }
            set { ExcelShortTopsheet = value; }
        }
        public Excel.Worksheet AllItemsSheet
        {
            get { return ExcelAllItemsSheet; }
            set { ExcelAllItemsSheet = value; }
        }
        public void openExcel()
        {
            myExcelApplication = null;

            myExcelApplication = new Excel.Application(); // create Excell App
            myExcelApplication.DisplayAlerts = false; // turn off alerts


            myExcelWorkbook = (Excel.Workbook)(myExcelApplication.Workbooks._Open(excelFilePath, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value)); // open the existing excel file

            int numberOfWorkbooks = myExcelApplication.Workbooks.Count; // get number of workbooks (optional)

            ExcelShortageSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[1]; // this worksheet has current day shortage sheet
            ExcelShortTopsheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[2]; // this worksheet has shortage with Top item besides it
            ExcelAllItemsSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[3]; // this worksheet has master shortage
            
            //myExcelWorkSheet.Name = "Shortage"; // define a name for the worksheet (optinal)

            worksheetCount = myExcelWorkbook.Worksheets.Count; // get number of worksheets
        }

        public void saveExcel()
        {
            try
            {
                myExcelWorkbook.SaveAs(excelFilePath, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value); // Save data in excel


                myExcelWorkbook.Close(true, excelFilePath, System.Reflection.Missing.Value); // close the worksheet
                
                //Auto open the excel that was saved
                System.Diagnostics.Process.Start(excelFilePath);

            }
            catch { }
        }
    }
}
