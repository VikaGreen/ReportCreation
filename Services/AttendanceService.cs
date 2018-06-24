using Microsoft.Office.Interop.Word;
using ReportCreation_2._0.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReportCreation_2._0.Services
{
    public class AttendanceService
    {
        public void SaveWord(AttendanceModel att, string filename)
        {
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
            
            document.Paragraphs.Add();
            document.Paragraphs[1].Range.Text = "ВОРОНЕЖСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ";
            document.Paragraphs[1].Range.Bold = 1;
            document.Paragraphs[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            document.Paragraphs[1].Range.Font.Size = 12;
            document.Paragraphs[1].Range.Font.Name = "Times New Roman";
            document.Paragraphs.Space1(); 
            
            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Font.Size = 12;
            document.Paragraphs[document.Paragraphs.Count].Range.Bold = 1;
            document.Paragraphs[document.Paragraphs.Count].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Лист учета посещаемости и текущей успеваемости обучающихся\r\n";

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Font.Size = 12;
            document.Paragraphs[document.Paragraphs.Count].Range.Bold = 0;
            document.Paragraphs[document.Paragraphs.Count].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Учебный год: " + att.year.ToString() + "/" + (att.year + 1).ToString() + " Семестр: " + att.semester;

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "ВГУ / филиал: " + att.branch;

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Факультет: " + att.department;

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Уровень образования: " + att.level;

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Специальность / направление: " + att.speciality;

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Курс: " + att.course.ToString() + " Группа: " + att.group.ToString();

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Дисциплина: " + att.subject;

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Вид занятия: ";

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.Text = "Преподаватель: " + att.teacher + "\r\n";

            document.Paragraphs.Add();
            document.Paragraphs[document.Paragraphs.Count].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            var colCount = att.students[1].records.Count + 3;
            var rowsCount = att.students.Count+2;
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
            table.Cell(1, 1).Range.Text = "№";
            table.Cell(1, 2).Range.Text = "Фамилия, Имя, Отчество";
            table.Cell(1, 3).Range.Text = "Дата";
            table.Cell(1, colCount).Range.Text = "Примечание";

            //объединение ячеек
            table.Cell(1, 1).Merge(table.Cell(2, 1));
            table.Cell(1, 2).Merge(table.Cell(2, 2));
            table.Cell(1, 3).Merge(table.Cell(1, colCount - 1));
            table.Cell(1, colCount-att.students[1].records.Count+1).Merge(table.Cell(2, colCount));

            //нумерация студентов
            for (int i=3;i<=rowsCount; i++)
            {
                table.Cell(i, 1).Range.Text = (i-2).ToString();
            }

            //фамилии студентов
            var row = 3;
            foreach (var student in att.students)
            {
                table.Cell(row, 2).Range.Text = student.FIO;
                row++;
            }

            //даты
            var col = 3;            
            foreach (var record in att.students.First().records)
            { 
                table.Cell(2, col).Range.Text = record.date;
                col++;
            }

            var rown = 3;
            
            //посещения
            foreach (var student in att.students)
            {
                var coln = 3;
                foreach (var record in student.records)
                {
                    table.Cell(rown, coln).Range.Text = record.note;
                    coln++;
                }
                
                rown++;
            }

            //примечания
            var coli = 3;
            foreach (var student in att.students)
            {
                table.Cell(coli, colCount).Range.Text = student.note;
                coli++;
            }
            
            document.SaveAs(filename);
            document.Close();
            word.Quit();
        }
    }
}