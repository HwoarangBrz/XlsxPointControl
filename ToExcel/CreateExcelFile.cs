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
            var rngTable = ws.Range(2, _col, 2, _col + 4);
            rngTable.Style.Font.Bold = true;
            rngTable.Style.Fill.BackgroundColor = XLColor.Aqua;
            rngTable.Merge();
            _col += 4;
        }

        private static void CreateWeekendDay(DateTime dateTime, IXLWorksheet ws)
        {
        }

        private static void CreateHoliday(DateTime dateTime, IXLWorksheet ws)
        {

        }
    }
}
