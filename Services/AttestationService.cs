using Microsoft.Office.Interop.Word;
using ReportCreation_2._0.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReportCreation_2._0.Services
{
    public class GroupResults
    {
        public int course { get; set; }
        public int group { get; set; }
        public string date { get; set; }
        public string contingent { get; set; }

        public int countOfStudens { get; set; }
        public float percentOfStudents { get; set; }

        public int countNA { get; set; }
        public int countA { get; set; }
        public int countB { get; set; }
        public int countC { get; set; }
        public int countD { get; set; }

        public float percentNA { get; set; }
        public float percentA { get; set; }
        public float percentB { get; set; }
        public float percentC { get; set; }
        public float percentD { get; set; }

        public float avg { get; set; }
    }

    public class Total
    {
        public int totalCount { get; set; }
        public int totalCountNA { get; set; }
        public float totalPercent { get; set; }
        public int totalCountA { get; set; }
        public int totalCountB { get; set; }
        public int totalCountC { get; set; }
        public int totalCountD { get; set; }
        public float totalPercentA { get; set; }
        public float totalPercentB { get; set; }
        public float totalPercentC { get; set; }
        public float totalPercentD { get; set; }
        public float totalAVG { get; set; }
    }

        public class AttestationService
    {
        public List<GroupResults> Calculation(AttestationModel att)
        {
            List<GroupResults> listGroups = new List<GroupResults>();
            int cnt = att.attestationRecords.Count;
            foreach (var record in att.attestationRecords)
            {
                GroupResults res = new GroupResults();
                res.course = record.course;
                res.group = record.group;
                res.date = record.date;
                res.contingent = record.contingentOfStudents;
                res.countOfStudens = record.marks.Count;
                int sum = 0;
                foreach (var mark in record.marks)
                {
                    switch (mark.mark)
                    {
                        case 5: res.countA++; sum = sum + 5;
                            break;
                        case 4: res.countB++; sum = sum + 4;
                            break;
                        case 3: res.countC++; sum = sum + 3;
                            break;
                        case 2: res.countD++; sum = sum + 2;
                            break;
                        default: res.countNA++; 
                            break;
                    }
                }
                                
                res.percentNA = (float)(res.countNA) / res.countOfStudens * 100;
                res.percentOfStudents = 100 - res.percentNA;
                res.percentA = (float)(res.countA) / res.countOfStudens * 100;
                res.percentB = (float)(res.countB) / res.countOfStudens * 100;
                res.percentC = (float)(res.countC) / res.countOfStudens * 100;
                res.percentD = (float)(res.countD) / res.countOfStudens * 100;
                res.avg = (float)(sum) / (res.countOfStudens - res.countNA);
                listGroups.Add(res);
            }
            return listGroups;
        }

        public Total TotalResult (List<GroupResults> list)
        {
            Total total = new Total();
            total.totalCount = 0;
            total.totalCountNA = 0;
            total.totalCountA = 0;
            total.totalCountB = 0;
            total.totalCountC = 0;
            total.totalCountD = 0;
            total.totalAVG = 0;
            foreach (var res in list)
            {
                total.totalCount = total.totalCount + res.countOfStudens;
                total.totalCountNA = total.totalCountNA + res.countNA;
                total.totalCountA = total.totalCountA + res.countA;
                total.totalCountB = total.totalCountB + res.countB;
                total.totalCountC = total.totalCountC + res.countC;
                total.totalCountD = total.totalCountD + res.countD;
                total.totalAVG = total.totalAVG + res.avg;

            }
            total.totalAVG = total.totalAVG / list.Count;
            int totalAnswered = total.totalCount - total.totalCountNA;
            foreach (var res in list)
            {
                total.totalPercent = (float)totalAnswered / total.totalCount * 100;
                total.totalPercentA = (float)total.totalCountA / totalAnswered * 100;
                total.totalPercentB = (float)total.totalCountB / totalAnswered * 100;
                total.totalPercentC = (float)total.totalCountC / totalAnswered * 100;
                total.totalPercentD = (float)total.totalCountD / totalAnswered * 100;

            }

            return total;
        }
        public void SaveWord(AttestationModel att, string filename)
        {
            List<GroupResults> listResults = Calculation(att);
            Total total = TotalResult(listResults);

            Microsoft.Office.Interop.Word._Application word = new Microsoft.Office.Interop.Word.Application();
            var document = word.Documents.Add();
            var fitBehavior = Type.Missing;

            document.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            document.PageSetup.LeftMargin = 28;
            document.PageSetup.TopMargin = 28;
            document.PageSetup.BottomMargin = 28;
            document.PageSetup.RightMargin = 28;
            word.ActiveDocument.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "Б ВГУ 170.750 – 2015";
            word.ActiveDocument.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            word.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Name = "Arial";
            word.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Size = 10;
            word.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Bold = 1;

            document.Paragraphs.Add();
            document.Paragraphs[1].Range.Text = "УТВЕРЖДАЮ";
            document.Paragraphs[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            document.Paragraphs[1].Range.Font.Name = "Arial";
            document.Paragraphs[1].Range.Font.Size = 12;
            document.Paragraphs.Space1();
            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "заведующий кафедрой";
            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "_____________________";
            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "__.__.20__\r\n";

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Bold = 1;
            document.Paragraphs[document.Paragraphs.Count].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Итоговые данные результатов промежуточной аттестации";

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Bold = 0;
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Специальность";

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Underline = WdUnderline.wdUnderlineSingle;
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Кафедра " + att.speciality;

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Underline = WdUnderline.wdUnderlineNone;
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Учебный год " + att.year.ToString() + "/" + (att.year + 1).ToString() + "\r\n";


            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            var colCount = 15;
            var rowsCount = 4 + listResults.Count;
            document.Tables.Add(document.Paragraphs[document.Paragraphs.Count].Range, rowsCount, colCount, ref fitBehavior,
                ref fitBehavior);
            var borders = new Border[6];
            var table = document.Tables[1];
            borders[0] = table.Borders[WdBorderType.wdBorderLeft];
            borders[1] = table.Borders[WdBorderType.wdBorderRight];
            borders[2] = table.Borders[WdBorderType.wdBorderTop];
            borders[3] = table.Borders[WdBorderType.wdBorderBottom];
            borders[4] = table.Borders[WdBorderType.wdBorderHorizontal];
            borders[5] = table.Borders[WdBorderType.wdBorderVertical];
            foreach (var border in borders)
            {
                border.LineStyle = WdLineStyle.wdLineStyleSingle;
                border.Color = WdColor.wdColorBlack;
            }
            table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
            table.Rows.HeightRule = WdRowHeightRule.wdRowHeightAuto;

            //объединение ячеек
            table.Cell(1, 5).Merge(table.Cell(1, colCount));
            table.Cell(2, 5).Merge(table.Cell(2, 6));
            table.Cell(2, 6).Merge(table.Cell(2, 7));
            table.Cell(2, 7).Merge(table.Cell(2, 8));
            table.Cell(2, 8).Merge(table.Cell(2, 9));
            table.Cell(2, 9).Merge(table.Cell(2, 10));
            table.Cell(1, 1).Merge(table.Cell(3, 1));
            table.Cell(1, 2).Merge(table.Cell(3, 2));
            table.Cell(1, 3).Merge(table.Cell(3, 3));
            table.Cell(1, 4).Merge(table.Cell(3, 4));
            table.Cell(rowsCount, 1).Merge(table.Cell(rowsCount, 3));
                     
            table.Cell(1, 1).Range.Text = "Наименование дисциплины";
            table.Cell(1, 2).Range.Text = "Дата проведения аттестации";
            table.Cell(1, 2).Range.Orientation = WdTextOrientation.wdTextOrientationUpward;            
            table.Cell(1, 3).Range.Text = "Группа, курс";
            table.Cell(1, 3).Range.Orientation = WdTextOrientation.wdTextOrientationUpward;
            table.Cell(1, 4).Range.Text = "Контингент студентов";
            table.Cell(1, 4).Range.Orientation = WdTextOrientation.wdTextOrientationUpward;
            table.Cell(1, 5).Range.Text = "Результаты текущей аттестации";
            table.Cell(2, 5).Range.Text = "Количество опрошенных";
            table.Cell(2, 6).Range.Text = "«Отл.»";
            table.Cell(2, 7).Range.Text = "«Хор.»";
            table.Cell(2, 8).Range.Text = "«Удовл.»";
            table.Cell(2, 9).Range.Text = "«Неуд.»";
            table.Cell(2, 10).Range.Text = "Средний балл";
            table.Cell(rowsCount, 1).Range.Text = "Итого по дисциплине";
            for (int i = 5; i<colCount;i++)
            {
                table.Cell(3, i).Range.Text = "абс.";
                i++;
            }
            for (int i = 6; i < colCount; i++)
            {
                table.Cell(3, i).Range.Text = "%";
                i++;
            }
            for (int i = 4; i<rowsCount;i++)
            {
                table.Cell(i, 1).Range.Text = att.subject;
            }
            var j = 4;
            foreach (var result in listResults)
            {
                table.Cell(j, 2).Range.Text = result.date;
                table.Cell(j, 3).Range.Text = result.group.ToString() + " г., " + result.course.ToString() + " курс";
                table.Cell(j, 4).Range.Text = result.contingent;
                table.Cell(j, 5).Range.Text = (result.countOfStudens - result.countNA).ToString();
                table.Cell(j, 6).Range.Text = (result.percentOfStudents).ToString();
                table.Cell(j, 7).Range.Text = (result.countA).ToString();
                table.Cell(j, 8).Range.Text = (result.percentA).ToString();
                table.Cell(j, 9).Range.Text = (result.countB).ToString();
                table.Cell(j, 10).Range.Text = (result.percentB).ToString();
                table.Cell(j, 11).Range.Text = (result.countC).ToString();
                table.Cell(j, 12).Range.Text = (result.percentC).ToString();
                table.Cell(j, 13).Range.Text = (result.countD).ToString();
                table.Cell(j, 14).Range.Text = (result.percentD).ToString();
                table.Cell(j, 15).Range.Text = (result.avg).ToString();
                j++;
            }

            table.Cell(rowsCount, 3).Range.Text = (total.totalCount - total.totalCountNA).ToString();
            table.Cell(rowsCount, 4).Range.Text = total.totalPercent.ToString();
            table.Cell(rowsCount, 5).Range.Text = total.totalCountA.ToString();
            table.Cell(rowsCount, 6).Range.Text = total.totalPercentA.ToString();
            table.Cell(rowsCount, 7).Range.Text = total.totalCountB.ToString();
            table.Cell(rowsCount, 8).Range.Text = total.totalPercentB.ToString();
            table.Cell(rowsCount, 9).Range.Text = total.totalCountC.ToString();
            table.Cell(rowsCount, 10).Range.Text = total.totalPercentC.ToString();
            table.Cell(rowsCount, 11).Range.Text = total.totalCountD.ToString();
            table.Cell(rowsCount, 12).Range.Text = total.totalPercentD.ToString();
            table.Cell(rowsCount, 13).Range.Text = total.totalAVG.ToString();

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Фамилии студентов, не явившихся на промежуточную аттестацию по неуважительной причине:";

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Ответственный исполнитель ____________________________________________";
 
            document.SaveAs(filename);
            document.Close();
            word.Quit();
        }
    }
}