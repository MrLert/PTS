using System;
using System.Reflection;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace PTS
{
    class WordMainReport : OfficeWord
    {
        private Application wordapp;
        private Document document;
        private Selection selection;
        private Style style;
        private int count;
        private Title infoTitle;
        private СarriageСharacteristic carriage;
        private List<Сharacteristics> characteristics;
        public WordMainReport(Title infoTitle, СarriageСharacteristic carriage, List<Сharacteristics> characteristics,int count)
        {
            wordapp = new Application() {Visible = true};
            document = wordapp.Documents.Add(Type.Missing, false, WdNewDocumentType.wdNewBlankDocument, true);
            document.SaveAs(@"C:\Users\Дмитрий\Documents\Универ\Диплом\PTS\PTS\"+ infoTitle.numberScheme+".doc", WdSaveFormat.wdFormatDocument);
            document.Content.Font.Size = 12;
            document.Content.Font.Name = "Times New Roman";
            this.infoTitle = infoTitle;
            this.carriage = carriage;
            this.characteristics = characteristics;
            this.count = count;
            document.Application.Selection.PageSetup.LeftMargin = document.Content.Application.CentimetersToPoints(2);
            document.Application.Selection.PageSetup.RightMargin = document.Content.Application.CentimetersToPoints(1);
            document.Application.Selection.PageSetup.TopMargin = document.Content.Application.CentimetersToPoints((float)1.5);
            document.Application.Selection.PageSetup.BottomMargin = document.Content.Application.CentimetersToPoints((float)3.2);
            document.Application.Selection.ParagraphFormat.LineSpacing = document.Content.Application.CentimetersToPoints((float)0.5);
            document.Application.Selection.ParagraphFormat.FirstLineIndent = document.Content.Application.CentimetersToPoints((float)1.25);
            document.Application.Selection.ParagraphFormat.SpaceAfter = 0;
        }

        //требуется проверка текста
        private void InputCarriageCharacteristic()  
        {
            var paragraph = document.Paragraphs.Last;
            string input;
            if (carriage.type == "Платформа")
            {
                input = "Характеристика платформы.";
                inputText(input,WdParagraphAlignment.wdAlignParagraphLeft, 2, 12);
                emptyline(1);
                input = "Для перевозки груза используем платформу с деревянным " +
                        "или деревометаллическим настилом пола (с шириной металлической полосы max 1300 мм), грузоподъемностью 63 - 71 т. Длина базы  - " +
                        carriage.baseLength.ToString() + "мм.Высота пола платформы от УГР -" + carriage.heightFromFloor +
                        "мм.";
                inputText(input,WdParagraphAlignment.wdAlignParagraphLeft, 0, 12);
            }
            else if (carriage.type == "Полувагон")
            {
                input = "Характеристика полувагона.";
                inputText(input, WdParagraphAlignment.wdAlignParagraphLeft, 2, 12);
                emptyline(1);
                input = "Для перевозки груза используем полувагон, грузоподъемностью 63 - 71 т. Длина базы  - " +
                        carriage.baseLength.ToString() + "мм.Высота пола полувагона от УГР -" + carriage.heightFromFloor +
                        "мм.";
                inputText(input, WdParagraphAlignment.wdAlignParagraphLeft, 0, 12);
            }
            else
            {
                input = "Характеристика крытого полувагона.";
                inputText(input, WdParagraphAlignment.wdAlignParagraphLeft, 2,12);
                emptyline(1);
                input = "Для перевозки груза используем крытого полувагона, грузоподъемностью 63 - 71 т. Длина базы  - " +
                        carriage.baseLength.ToString() + "мм.Высота пола крытого полувагона от УГР -" + carriage.heightFromFloor +
                        "мм.";
                inputText(input, WdParagraphAlignment.wdAlignParagraphLeft, 0,12);
            }
        }

        private Paragraph emptyline(int n)
        {
            var paragraph = document.Paragraphs.Last;
            for (var i = 0; i < n; i++)
            {
                paragraph.Range.InsertParagraphAfter();
                paragraph = document.Paragraphs.Last;
            }
            return paragraph;
        }

        private void inputText(string input="", WdParagraphAlignment param=WdParagraphAlignment.wdAlignParagraphLeft, int bold=0,int size=12, int italic = 0, WdUnderline underline = WdUnderline.wdUnderlineNone)
        {
            var paragraph = emptyline(1);
            paragraph.Format.Alignment = param;
            paragraph.Range.Font.Bold = bold;
            paragraph.Range.Text = input;
            paragraph.Range.Font.Size = size;
            paragraph.Range.Font.Italic = italic;
            paragraph.Range.Font.Underline = underline;
        }

        public void createTitle()
        {
            var paragraphs = document.Paragraphs;
            var paragraph = paragraphs[1];
            paragraph.Range.Text = "УТВЕРЖДАЮ:" + Environment.NewLine + 
                infoTitle.post + Environment.NewLine + 
                infoTitle.organization + Environment.NewLine +
                "_______________/" + infoTitle.nameSender + "/" + Environment.NewLine +
                "\"___\" ________2017г.";
            paragraph = emptyline(14);
            for (var i = 1; i < 6; i++)
            {
                paragraphs[i].Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;    
            }
            paragraph.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.Range.Font.Size = 14;
            paragraph.Range.Font.Bold = 2;
            paragraph.Range.Text = "Расчетно-пояснительная записка";
            paragraph = emptyline(1);
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Bold = 0;
            paragraph.Range.Text = "к схеме размещения и крепления груза" + Environment.NewLine;
            paragraph.Range.Text = infoTitle.nameFreight + Environment.NewLine;
            if (carriage.type == "Платформа")
            {
                paragraph.Range.Text = "на ж.д. платформе." + Environment.NewLine;
            }
            else if(carriage.type == "Полувагон")
            {
                paragraph.Range.Text = "в ж.д. полувагоне." + Environment.NewLine;
            }
            else
            {
                paragraph.Range.Text = "в ж.д. крытом полувагоне" + Environment.NewLine;
            }
            paragraph = emptyline(19);
            paragraph.Range.Text = "Расчет произвел   ___________________      " + infoTitle.nameDeveloper;
            Characteristic();
        }

        private void createTable()
        {
            var paragraph = document.Paragraphs.Last;
            paragraph.Range.Font.Bold = 0;
            paragraph.Range.Font.Size = 10;
            var table = document.Tables.Add(document.Paragraphs.Last.Range,count+2, 9,WdDefaultTableBehavior.wdWord9TableBehavior,WdAutoFitBehavior.wdAutoFitWindow);
            table.set_Style("Сетка таблицы");
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    if (cell.RowIndex == 1)
                    {
                        switch (cell.ColumnIndex)
                        {
                            case 1:
                                cell.Range.Text = "№";
                                break;
                            case 2:
                                cell.Range.Text = "Наименование груза";
                                break;
                            case 3:
                                cell.Range.Text = "Масса, кг";
                                break;
                            case 4:
                                cell.Range.Text = "Длина, мм";
                                break;
                            case 5:
                                cell.Range.Text = "Ширина, мм";
                                break;
                            case 6:
                                cell.Range.Text = "Высота, мм";
                                break;
                            case 7:
                                cell.Range.Text = "Высота ЦТ от основания груза, мм";
                                break;
                            case 8:
                                cell.Range.Text =
                                    "Наименьшее расстояние от проекции ЦТ на плоскость опорной поверхности до линии опрокидывания, мм";
                                break;
                        }
                    }
                    else if (cell.RowIndex == 2)
                    {
                        switch (cell.ColumnIndex)
                        {
                            case 8:
                                cell.Range.Text = "вдоль вагона (Lпр )";
                                break;
                            case 9:
                                cell.Range.Text = "поперек вагона (Bп)";
                                break;
                        }
                    }
                    else
                    {
                        cell.Range.Font.Size = 12;
                        switch (cell.ColumnIndex)
                        {
                            case 1:
                                cell.Range.Text = characteristics[cell.RowIndex - 3].number.ToString();
                                break;
                            case 2:
                                cell.Range.Text = characteristics[cell.RowIndex - 3].name;
                                break;
                            case 3:
                                cell.Range.Text = characteristics[cell.RowIndex - 3].weight.ToString();
                                break;
                            case 4:
                                cell.Range.Text = characteristics[cell.RowIndex - 3].length.ToString();
                                break;
                            case 5:
                                cell.Range.Text = characteristics[cell.RowIndex - 3].width.ToString();
                                break;
                            case 6:
                                cell.Range.Text = characteristics[cell.RowIndex - 3].height.ToString();
                                break;
                            case 7:
                                cell.Range.Text = characteristics[cell.RowIndex - 3].centerOfGravity.ToString();
                                break;
                            case 8:
                                cell.Range.Text = characteristics[cell.RowIndex - 3].Lpr.ToString();
                                break;
                            case 9:
                                cell.Range.Text = characteristics[cell.RowIndex - 3].Bp.ToString();
                                break;
                        }
                    }
                }
            }
            for (var i = 1; i < 8; i++)
            {
                table.Cell(1,i).Merge(table.Cell(2,i));
            }
            table.Cell(1,8).Merge(table.Cell(1,9));
        }

        public void Characteristic()
        {
            var paragraph = emptyline(2);
            inputText("Задача расчёта.", WdParagraphAlignment.wdAlignParagraphCenter,2, 12);
            paragraph = emptyline(2);
            var input = "Задачей расчета является размещение и определение прочности крепления грузов ";
            if (carriage.type == "Платформа")
            {
                input += "на железнодорожной платформе ";
            }
            else if (carriage.type == "Полувагон")
            {
                input += "в железнодорожном полувагоне ";
            }
            else
            {
                input += "в железнодорожном крытом полувагоне";
            }
            input += "от продольных и поперечных смещений при скорости движения поезда 100 км/ч.";
            inputText(input, WdParagraphAlignment.wdAlignParagraphLeft,0, 12);
            paragraph = emptyline(2);
            inputText("Характеристика груза.", WdParagraphAlignment.wdAlignParagraphLeft, 2, 12);
            paragraph = emptyline(2);
            createTable();
            paragraph = emptyline(2);
            InputCarriageCharacteristic();
            centerGravity();
        }

        private void cancelUnderline(int symbol)
        {
            wordapp.Selection.EndKey(WdUnits.wdStory,WdMovementType.wdMove);
            wordapp.Selection.HomeKey(WdUnits.wdLine, WdMovementType.wdMove);
            wordapp.Selection.MoveRight(WdUnits.wdCharacter, symbol, WdMovementType.wdExtend);
            wordapp.Selection.Font.Underline = WdUnderline.wdUnderlineNone;
            wordapp.Selection.EndKey(WdUnits.wdStory, WdMovementType.wdMove);
        }

        private void centerGravity()
        {
            var paragraph = emptyline(2);
            var input = "Определение высоты общего центра тяжести грузов в вагоне над уровнем головки рельса:";
            paragraph.Format.FirstLineIndent = document.Content.Application.CentimetersToPoints(0);
            inputText(input, WdParagraphAlignment.wdAlignParagraphLeft, 2);
            input = "(стр.45, формула 19)";
            inputText(input, WdParagraphAlignment.wdAlignParagraphRight, 0, 10, 0, WdUnderline.wdUnderlineSingle);

            input = "Hцт.гро=Qгр1*Hцт1+Qгр2*Нцт2+...+QгрN*НцтN";
            inputText(input, WdParagraphAlignment.wdAlignParagraphLeft,0,12,1, WdUnderline.wdUnderlineSingle);
            cancelUnderline(8);
            input = "                           Qгр1+Qгр2+...+QгрN";
            inputText(input, WdParagraphAlignment.wdAlignParagraphLeft, 0, 12, 1);
            paragraph = emptyline(1);

            for (var i = 1; i <= count; i++)
            {
                input = "Нцт" + i + "=" +carriage.heightFromFloor+"+" +characteristics[i - 1].centerOfGravity+"+" +
                        characteristics[i - 1].heightAboveFloor+"="+ (carriage.heightFromFloor + characteristics[i - 1].centerOfGravity + characteristics[i - 1].heightAboveFloor)+" мм";
                inputText(input);
            }
            paragraph = emptyline(1);
            input = "Hцт.гро= ";
            for (var i = 1; i <= count; i++)
            {
                if (i != 1)
                    input += "+";
                input += characteristics[i - 1].weight / 1000 + "*" +
                         (carriage.heightFromFloor + characteristics[i - 1].centerOfGravity +
                          characteristics[i - 1].heightAboveFloor);
            }
            inputText(input, WdParagraphAlignment.wdAlignParagraphLeft, 0, 12, 0, WdUnderline.wdUnderlineSingle);
            cancelUnderline(8);

            CloseAndSave();
        }

        public void CloseAndSave()
        {
            wordapp.Quit(WdSaveOptions.wdSaveChanges, WdOriginalFormat.wdWordDocument);
            wordapp = null;
        }
    }
}