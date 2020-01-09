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

            ws.Column("A").Width = 0.92;
            ws.Row(1).Height = 8.25;

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
            wb.ReferenceStyle = XLReferenceStyle.R1C1;
            wb.CalculateMode = XLCalculateMode.Auto;
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ControleDePonto_" + _reference.Year + "_" + _reference.Month + ".xlsx";
            wb.SaveAs(filePath);
            Console.WriteLine("Concluído. Verifique em " + filePath);
        }

        private static void CreateUtilDay(DateTime dateTime, IXLWorksheet ws)
        {
            ws.Cell(2, _col).Value = culture.DateTimeFormat.GetDayName(dateTime.DayOfWeek);
            ws.Range(2, _col, 2, _col + 3).FirstRow().Merge();
            ws.Cell(3, _col).Value = dateTime.ToString("dd/MMM/yyyy");
            ws.Range(3, _col, 3, _col + 2).FirstRow().Merge();
            ws.Cell(4, _col).Value     = "Tarefa";
            ws.Cell(4, _col + 1).Value = "Início";
            ws.Cell(4, _col + 2).Value = "Final";
            string cell = GetExcelColumnName(_col + 3);
            ws.Cell(cell + (_col + 3)).FormulaA1 = "=SUM(" + cell + "5:$" + cell + "$14)";
            ws.Range(3, _col + 3, 4, _col + 3).Style.NumberFormat.Format = "hh:mm";
            var rngHeader = ws.Range(2, _col, 4, _col + 3);
            rngHeader.Style.Font.SetBold().Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            rngHeader.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rngHeader.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            for (int i = 0; i < 4; i++)
            {
                var col = ws.Column(_col + i);
                col.Width = 8.43;
            }
            var rngBody = ws.Range(5, _col, 14, _col + 3);
            rngBody.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            rngBody.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rngBody.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            var rngData = ws.Range(5, _col + 1, 14, _col + 3);
            rngData.Style.NumberFormat.Format = "hh:mm";
            _col += 4;
        }

        private static void CreateWeekendDay(DateTime dateTime, IXLWorksheet ws)
        {
            ws.Cell(2, _col).Value = culture.DateTimeFormat.GetDayName(dateTime.DayOfWeek);
            ws.Cell(3, _col).Value = dateTime.ToString("dd/MMM/yyyy");
            ws.Column( _col).Width = 15;
            var rngHeader = ws.Range(2, _col, 4, _col);
            rngHeader.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rngHeader.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            rngHeader.Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.DarkCyan).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            var rngBody = ws.Range(5, _col, 14, _col);
            rngBody.Style.Fill.SetBackgroundColor(XLColor.Cornsilk).Border.OutsideBorder = XLBorderStyleValues.Thick;
            _col += 1;
        }

        private static void CreateHoliday(DateTime dateTime, IXLWorksheet ws)
        {
            ws.Cell(2, _col).Value = culture.DateTimeFormat.GetDayName(dateTime.DayOfWeek);
            ws.Cell(3, _col).Value = dateTime.ToString("dd/MMM/yyyy");
            ws.Cell(4, _col).Value = Holiday.GetHolidayName(dateTime);
            ws.Column(_col).Width = 15;
            var rngHeader = ws.Range(2, _col, 4, _col);
            rngHeader.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rngHeader.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            rngHeader.Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.DarkCyan).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            var rngBody = ws.Range(5, _col, 14, _col);
            rngBody.Style.Fill.SetBackgroundColor(XLColor.Cornsilk).Border.OutsideBorder = XLBorderStyleValues.Thick;
            _col += 1;
        }
        
        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
