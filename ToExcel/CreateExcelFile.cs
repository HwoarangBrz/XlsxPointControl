using System;
using System.Globalization;
using ClosedXML.Excel;

namespace ToExcel
{
    public static class CreateExcelFile
    {
        private static DateTime _reference;
        private static int _col;
        private static CultureInfo culture = new CultureInfo("pt-BR");

        static CreateExcelFile()
        {
        }

        public static void Dummy()
        {
            Console.WriteLine("Dummy");
        }

        public static void GenerateWorksheet(DateTime dateTime)
        {
            _reference = dateTime;
            _col = 2;
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add(culture.DateTimeFormat.GetMonthName(dateTime.Month) + "." + dateTime.Year);
            int firstDayOfMonth = new DateTime(_reference.Year, _reference.Month, 1).Day;
            int lastDayOfMonth = _reference.AddMonths(1).AddDays(-1).Day;

            for (int i = firstDayOfMonth; i <= lastDayOfMonth; i++)
            {
                DateTime actualDate = _reference.AddDays(i - 1);
                if (actualDate.DayOfWeek == DayOfWeek.Saturday || actualDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    CreateWeekendDay(actualDate, ws);
                }
                else
                {
                    if (Holiday.IsHolidayDate(actualDate))
                    {
                        CreateHoliday(actualDate, ws);
                    }
                    else
                    {
                        CreateUtilDay(actualDate, ws);
                    }
                }
            }
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ControleDePonto_" + _reference.Year + "_" + _reference.Month + ".xlsx";
            wb.SaveAs(filePath);
            Console.WriteLine("Concluído. Verifique em " + filePath);
        }

        private static void CreateUtilDay(DateTime dateTime, IXLWorksheet ws)
        {
            ws.Cell(2, _col).Value = culture.DateTimeFormat.GetDayName(dateTime.DayOfWeek);
            var rngHeader_l1 = ws.Range(2, _col, 2, _col + 3);
            rngHeader_l1.FirstCell().Style.Font.SetBold().Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            rngHeader_l1.FirstRow().Merge();
            ws.Cell(3, _col).Value = dateTime.ToString("dd/MMM/yyyy");
            var rngHeader_l2 = ws.Range(3, _col, 3, _col + 2);
            rngHeader_l2.FirstCell().Style.Font.SetBold().Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            rngHeader_l2.FirstRow().Merge();
            ws.Cell(4, _col).Value     = "Tarefa";
            ws.Cell(4, _col + 1).Value = "Início";
            ws.Cell(4, _col + 2).Value = "Final";
            ws.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            ws.Columns(_col + "-" + _col + 3).AdjustToContents();
            for (int i = 0; i < 4; i++)
            {
                var col = ws.Column(_col + i);
                col.Width = 8.43;
            }
            
            _col += 4;
        }

        private static void CreateWeekendDay(DateTime dateTime, IXLWorksheet ws)
        {
            ws.Cell(2, _col).Value = culture.DateTimeFormat.GetDayName(dateTime.DayOfWeek);
            ws.Cell(2, _col).Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.DarkPowderBlue).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            ws.Cell(3, _col).Value = dateTime.ToString("dd/MMM/yyyy");
            ws.Cell(3, _col).Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.DarkPowderBlue).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            ws.Cell(4, _col).Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.DarkPowderBlue).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            ws.Column( _col).Width = 15;
            _col += 1;
        }

        private static void CreateHoliday(DateTime dateTime, IXLWorksheet ws)
        {

        }
    }
}
