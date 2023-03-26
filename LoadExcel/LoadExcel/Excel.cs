using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace LoadExcel
{
    public class Excel
    {
        public void WriteExcel()
        {
            string excelPath = "C:\\Users\\zzang\\Downloads\\test.xlsx";
            var excel = new Application();
           
            Workbook workBook = null;
            Worksheet workSheet = null;

            //open the workbook
            workBook = excel.Workbooks.Open(excelPath);
            // first worksheet if 2nd worksheet requried then we can change to 2
            workSheet = workBook.Worksheets[1];
            //range of excel 
            Range cell = workSheet.Range["B2:E2"];
            // inserting into excel as array values
            string[] cellValues = new[] { "test1", "test2", "test3", "test4" };
           // cell.Value2(XlRangeValueDataType.xlRangeValueDefault, cellValues);
            //cell.Value(XlRangeValueDataType.xlRangeValueDefault, cellValues);
            cell.set_Value(XlRangeValueDataType.xlRangeValueDefault, cellValues);
            
            try
            {
                workBook.SaveAs("C:\\Users\\zzang\\Downloads\\test.xlsx");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Did not save into Excel error message: {ex.Message}");

            }
            finally
            {
                workBook.Close();
            }


        }


    }
}
