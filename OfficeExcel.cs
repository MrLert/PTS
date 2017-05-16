using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;


namespace PTS
{
    public class OfficeExcel
    {
        private СarriageСharacteristic CarriageList;
        public int count;
        private List<Сharacteristics> CharacteristicsList;
        private Title infoTitle;
        private Application excel;
        private readonly Workbook wb;
        private Workbooks wbs;
        
        

        public OfficeExcel()
        {
            excel = new Application {Visible = false};
            wbs = excel.Workbooks;
            wb = excel.Workbooks.Open(@"C:\Users\Дмитрий\Documents\Универ\Диплом\PTS\PTS\1.xls");
            CarriageList = new СarriageСharacteristic();
            CharacteristicsList = new List<Сharacteristics>();
        }

        public СarriageСharacteristic InputListCarriage()
        {
            var excelsheets = wb.Worksheets;
            var excelsheet = excelsheets.get_Item(1);
            var excelcells = excelsheet.Range("B" + 4, Type.Missing);
            CarriageList.type = Convert.ToString(excelcells.Value2);
            excelcells = excelsheet.Range("C" + 4, Type.Missing);
            CarriageList.weight = Convert.ToDouble(excelcells.Value2);
            excelcells = excelsheet.Range("D" + 4, Type.Missing);
            CarriageList.baseLength = Convert.ToDouble(excelcells.Value2);
            excelcells = excelsheet.Range("E" + 4, Type.Missing);
            CarriageList.heightFromFloor = Convert.ToDouble(excelcells.Value2);
            excelcells = excelsheet.Range("G" + 4, Type.Missing);
            CarriageList.centerOfGravity = Convert.ToDouble(excelcells.Value2);
            excelcells = excelsheet.Range("I" + 4, Type.Missing);
            CarriageList.length = Convert.ToDouble(excelcells.Value2);
            excelcells = excelsheet.Range("K" + 4, Type.Missing);
            CarriageList.width = Convert.ToDouble(excelcells.Value2);
            excelcells = excelsheet.Range("L" + 4, Type.Missing);
            CarriageList.windwardSurfaceArea = Convert.ToDouble(excelcells.Value2);
            return CarriageList;
        }

        public List<Сharacteristics> InputListСharacteristicses()
        {
            var excelsheets = wb.Worksheets;
            var excelsheet = excelsheets.get_Item(1);
            var excelcells = excelsheet.Range("B1", Type.Missing);
            count = Convert.ToInt32(excelcells.Value2);
            
            for (int i = 0; i < count; i++)
            {
                Сharacteristics input = new Сharacteristics(); 
                var number = 7+i*2;
                input.number = i + 1;
                excelcells = excelsheet.Range("C" + number, Type.Missing);
                input.name = Convert.ToString(excelcells.Value2);
                excelcells = excelsheet.Range("D" + number, Type.Missing);
                input.weight = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("E" + number, Type.Missing);
                input.length = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("F" + number, Type.Missing);
                input.width = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("G" + number, Type.Missing);
                input.height = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("H" + number, Type.Missing);
                input.centerOfGravity = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("K" + number, Type.Missing);
                input.Lpr = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("L" + number, Type.Missing);
                input.Bp = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("M" + number, Type.Missing);
                input.L_CT = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("N" + number, Type.Missing);
                input.B_CT = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("O" + number, Type.Missing);
                input.coefficientOfFriction = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("T" + number, Type.Missing);
                input.heightOfLongitudinal = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("U" + number, Type.Missing);
                input.heightOfTransverse = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("V" + number, Type.Missing);
                input.windwardSurfaceArea = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("W" + number, Type.Missing);
                input.heightAboveFloor = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("X" + number, Type.Missing);
                input.HeightOfProtruding = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("Y" + number, Type.Missing);
                input.coefficientOfFrictionTransverse = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("Z" + number, Type.Missing);
                input.additionalLongitudinalLoad = Convert.ToDouble(excelcells.Value2);
                excelcells = excelsheet.Range("AA" + number, Type.Missing);
                input.additionalLateralLoad = Convert.ToDouble(excelcells.Value2);
                CharacteristicsList.Add(input);
            }
            return CharacteristicsList;
        }

        public Title InpuTitle()
        {
            var excelsheets = wb.Worksheets;
            var excelsheet = excelsheets.get_Item(3);
            var excelcells = excelsheet.Range("B" + 1, Type.Missing);
            infoTitle.post = Convert.ToString(excelcells.Value2);
            excelcells = excelsheet.Range("B" + 2, Type.Missing);
            infoTitle.organization = Convert.ToString(excelcells.Value2);
            excelcells = excelsheet.Range("B" + 3, Type.Missing);
            infoTitle.nameSender = Convert.ToString(excelcells.Value2);
            excelcells = excelsheet.Range("B" + 4, Type.Missing);
            infoTitle.nameFreight = Convert.ToString(excelcells.Value2);
            excelcells = excelsheet.Range("B" + 5, Type.Missing);
            infoTitle.nameDeveloper = Convert.ToString(excelcells.Value2);
            excelcells = excelsheet.Range("B" + 6, Type.Missing);
            infoTitle.numberScheme = Convert.ToString(excelcells.Value2);
            return infoTitle;
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