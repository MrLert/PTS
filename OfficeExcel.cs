using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
namespace PTS
{
    public class OfficeExcel
    {
        private Microsoft.Office.Interop.Excel.Application excel;
        private Workbooks wbs;
        private Workbook wb;
        public List<СarriageСharacteristic> CarriageList;

        public OfficeExcel()
        {
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            wbs = excel.Workbooks;
            wb = excel.Workbooks.Open(@"C:\Users\Дмитрий\Documents\Visual Studio 2017\Projects\PTS\PTS\1.xls");
            CarriageList = new List<СarriageСharacteristic>();
            GetInformation();
        }

        public void GetInformation()
        { 
            var excelsheets = wb.Worksheets;
            var excelsheet = excelsheets.get_Item(1);
            var excelcells = excelsheet.Range("B1", Type.Missing);
            var numberOfLoads = Convert.ToInt32(excelcells.Value2);
            СarriageСharacteristic inputCarriage = new СarriageСharacteristic();
            var letter = 'A';
            letter++;
            
            excelcells = excelsheet.Range(letter+"4", Type.Missing);
            inputCarriage.type = Convert.ToString(excelcells.Value2);
            CloseExcel();
        }
        public void CloseExcel()
        {
            excel.Workbooks.Close();
            excel.Quit();
            excel = null;
            GC.Collect();
        }
    }
}