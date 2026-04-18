using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using ReportGeneration_Klimov.Items;
using ReportGeneration_Klimov.Models;
using ReportGeneration_Klimov.Pages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportGeneration_Klimov.Classes.Common
{
    public class Report
    {
        public static void Group(int IdGroup, Main main)
        {
            SaveFileDialog SFD = new SaveFileDialog
            {
                InitialDirectory = @"C:\",
                Filter = "Excel (*.xlsx) | *.xlsx"
            };

            SFD.ShowDialog();

            if (SFD.FileName != "")
            {
                Group Group = Main.connection.Groups.ToList().Find(x => x.Id == IdGroup);

                var ExcelApp = new Excel.Application();

                try
                {
                    ExcelApp.Visible = false;
                    Excel.Workbook Workbook = ExcelApp.Workbooks.Add(Type.Missing);
                    Excel.Worksheet Worksheet = Workbook.ActiveSheet;

                    (Worksheet.Cells[1, 1] as Excel.Range).Value = $"Отчёт о группе {Group.Name}";
                    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 5]].Merge();
                    Styles(Worksheet.Cells[1, 1], 18);

                    (Worksheet.Cells[3, 1] as Excel.Range).Value = $"Список группы:";
                    Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3, 5]].Merge();
                    Styles(Worksheet.Cells[3, 1], 12, XlHAlign.xlHAlignLeft);

                    (Worksheet.Cells[4, 1] as Excel.Range).Value = $"ФИО";
                    Styles(Worksheet.Cells[4, 1], 12, Excel.XlHAlign.xlHAlignCenter, true);
                    (Worksheet.Cells[4, 1] as Excel.Range).ColumnWidth = 35.0f;

                    (Worksheet.Cells[4, 2] as Excel.Range).Value = $"Кол-во не сданных практических";
                    Styles(Worksheet.Cells[4, 2], 12, XlHAlign.xlHAlignCenter, true);

                    (Worksheet.Cells[4, 3] as Excel.Range).Value = $"Кол-во не сданных теоретических";
                    Styles(Worksheet.Cells[4, 3], 12, XlHAlign.xlHAlignCenter, true);

                    (Worksheet.Cells[4, 4] as Excel.Range).Value = $"Отсутствовал на паре";
                    Styles(Worksheet.Cells[4, 4], 12, XlHAlign.xlHAlignCenter, true);

                    (Worksheet.Cells[4, 5] as Excel.Range).Value = $"Опоздал";
                    Styles(Worksheet.Cells[4, 5], 12, XlHAlign.xlHAlignCenter, true);

                    int Height = 5;

                    int bestRow = -1;
                    int bestDebts = 999999;
                    int bestAbsences = 999999;

                    List<Models.Student> Students = Main.connection.Students.ToList().FindAll(x => x.IdGroup == IdGroup);
                    foreach (Models.Student Student in Students)
                    {
                        List<Discipline> StudentDisciplines = Main.connection.Disciplines.ToList().FindAll(
                            x => x.IdGroup == Student.IdGroup);
                        int PracticeCount = 0;
                        int TheoryCount = 0;
                        int AbsenteeismCount = 0;
                        int LateCount = 0;

                        foreach (Discipline StudentDiscipline in StudentDisciplines)
                        {
                            List<Work> StudentWorks = Main.connection.Works.ToList().FindAll(x => x.IdDiscipline == StudentDiscipline.Id);
                            foreach (Work StudentWork in StudentWorks)
                            {
                                Evaluation Evaluation = Main.connection.Evaluations.ToList().Find(x =>
                                    x.IdWork == StudentWork.Id &&
                                    x.IdStudent == Student.Id);

                                if ((Evaluation != null && (Evaluation.Value.Trim() == "" || Evaluation.Value.Trim() == "2"))
                                    || Evaluation == null)
                                {
                                    if (StudentWork.IdType == 1)
                                        PracticeCount++;
                                    else if (StudentWork.IdType == 2)
                                        TheoryCount++;
                                }

                                if (Evaluation != null && Evaluation.Lateness.Trim() != "")
                                {
                                    if (Convert.ToInt32(Evaluation.Lateness) == 90)
                                        AbsenteeismCount++;
                                    else
                                        LateCount++;
                                }

                                int totalDebts = PracticeCount + TheoryCount;
                                int totalAbsences = AbsenteeismCount + LateCount;

                                if (totalDebts < bestDebts || (totalDebts == bestDebts && totalAbsences <= bestAbsences))
                                {
                                    bestDebts = totalDebts;
                                    bestAbsences = totalAbsences;
                                    bestRow = Height;
                                }
                            }
                        }

                        (Worksheet.Cells[Height, 1] as Excel.Range).Value = $"{Student.LastName} {Student.FirstName}";
                        Styles(Worksheet.Cells[Height, 1], 12, XlHAlign.xlHAlignLeft, true);

                        (Worksheet.Cells[Height, 2] as Excel.Range).Value = PracticeCount.ToString();
                        Styles(Worksheet.Cells[Height, 2], 12, XlHAlign.xlHAlignCenter, true);

                        (Worksheet.Cells[Height, 3] as Excel.Range).Value = TheoryCount.ToString();
                        Styles(Worksheet.Cells[Height, 3], 12, XlHAlign.xlHAlignCenter, true);

                        (Worksheet.Cells[Height, 4] as Excel.Range).Value = AbsenteeismCount.ToString();
                        Styles(Worksheet.Cells[Height, 4], 12, XlHAlign.xlHAlignCenter, true);

                        (Worksheet.Cells[Height, 5] as Excel.Range).Value = LateCount.ToString();
                        Styles(Worksheet.Cells[Height, 5], 12, XlHAlign.xlHAlignCenter, true);

                        Height++;
                    }

                    if (bestRow != -1)
                    {
                        Worksheet.Range[$"A{bestRow}:E{bestRow}"].Interior.Color = 6551089;
                        Worksheet.Range[$"A{bestRow}:E{bestRow}"].Font.Bold = true;
                    }

                    Workbook.SaveAs(SFD.FileName);
                    Workbook.Close();

                }
                catch (Exception ex) { };

                ExcelApp.Quit();
            }
        }

        public static void Styles(Excel.Range Cell, int FontSize,
            Excel.XlHAlign Position = Excel.XlHAlign.xlHAlignCenter, bool Border = false)
        {
            Cell.Font.Name = "Bahnschrift Light Condensed";
            Cell.Font.Size = FontSize;
            Cell.HorizontalAlignment = Position;
            Cell.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

            if (Border)
            {
                Excel.Borders border = Cell.Borders;
                border.LineStyle = Excel.XlLineStyle.xlDouble;
                border.Weight = XlBorderWeight.xlThin;
                Cell.WrapText = true;
            }
        }
    }
}
