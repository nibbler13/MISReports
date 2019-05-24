using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
    class FssInfo : ExcelGeneral {
        public static bool Process(string resultFile) {
            if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
                out Excel.Worksheet ws))
                return false;

            int usedRows = ws.UsedRange.Rows.Count;

            ws.Range["A3:BC3"].Select();
            xlApp.Selection.Copy();
            ws.Range["A4:BC" + usedRows].Select();
            xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
            ws.Range["A1"].Select();

            try {
                FssInfoAddPivotTable(wb, ws, xlApp);
            } catch (Exception e) {
                Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
            }

            SaveAndCloseWorkbook(xlApp, wb, ws);

            return true;
        }

        public static void PerformData(ref DataTable dataTable) {
            Dictionary<string[], string> addresses = new Dictionary<string[], string> {
                { new string[] { "вадковский", "18" }, "г.Москва, Вадковский переулок, д.18" },
                { new string[] { "комсомольский", "28" }, "г.Москва, Комсомольский проспект, д.28" },
                { new string[] { "комсосольский", "28" }, "г.Москва, Комсомольский проспект, д.28" },
                { new string[] { "сущевский", "12" }, "г.Москва, ул. Сущевский Вал, д.12" },
                { new string[] { "ленинградское", "16" }, "г.Москва, ул. Ленинградское шоссе, д.16" },
                { new string[] { "последний", "28" }, "г.Москва, ул. Последний переулок, д.28" },
                { new string[] { "тульская", "10" }, "г.Москва, ул. Большая Тульская, д.10 с9" },
                { new string[] { "лесная", "41" }, "г.Москва, ул. Лесная, д.41" },
                { new string[] { "триумфальная", "12" }, "г.Сочи, ул. Триумфальная, д.12а" },
                { new string[] { "лиговский", "274" }, "г.Санкт-Петербург, Лиговский проспект, д.274А" },
                { new string[] { "тургенева", "96" }, "г.Краснодар, ул. Тургенева, д.96" },
                { new string[] { "октября", "6" }, "г.Уфа, проспект Октября, д.6/1" },
                { new string[] { "нариманова", "65" }, "г.Казань, ул. Нариманова, д.65" },
                { new string[] { "бажова", "3" }, "г.Каменск-Уральский, ул.Бажова, д.3" },
            };

            for (int r = 0; r < dataTable.Rows.Count; r++) {
                if (DateTime.TryParse(dataTable.Rows[r][0].ToString(), out DateTime date)) {
                    dataTable.Rows[r][1] = date.Year;
                    dataTable.Rows[r][2] = "Квартал " + GetQuarter(date);
                    dataTable.Rows[r][3] = date.ToString("MMMM", CultureInfo.CreateSpecificCulture("ru"));
                    int weekNumber = GetIso8601WeekOfYear(date);
                    dataTable.Rows[r][4] = "Неделя " + (weekNumber.ToString().Length < 2 ? "0" + weekNumber : weekNumber.ToString());
                }

                string address = dataTable.Rows[r][17].ToString().ToLower();
                if (!string.IsNullOrEmpty(address) && !string.IsNullOrWhiteSpace(address)) 
                    foreach (KeyValuePair<string[], string> pair in addresses) 
                        if (address.Contains(pair.Key[0]) && address.Contains(pair.Key[1]))
                            dataTable.Rows[r][17] = pair.Value;
            }
        }

        private static int GetIso8601WeekOfYear(DateTime time) {
            // Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll 
            // be the same week# as whatever Thursday, Friday or Saturday are,
            // and we always get those right
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday) {
                time = time.AddDays(3);
            }

            // Return the week of our adjusted day
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        private static int GetQuarter(DateTime date) {
            if (date.Month >= 1 && date.Month <= 3)
                return 1;
            else if (date.Month >= 4 && date.Month <= 6)
                return 2;
            else if (date.Month >= 7 && date.Month <= 9)
                return 3;
            else
                return 4;
        }

        private static void FssInfoAddPivotTable(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
            ws.Cells[3, 1].Select();

            string pivotTableName = @"FssInfoPivotTable";
            Excel.Worksheet wsPivote = wb.Sheets["Сводная Таблица"];

            int wsRowsUsed = ws.UsedRange.Rows.Count;

            wsPivote.Activate();

            Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.Range["A2:BG" + wsRowsUsed], 6);
            Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

            pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

            pivotTable.PivotFields("Год").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Год").Position = 1;
            pivotTable.PivotFields("Год").AutoSort(Excel.XlSortOrder.xlDescending, "Год");

            pivotTable.PivotFields("Квартал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Квартал").Position = 2;
            pivotTable.PivotFields("Квартал").AutoSort(Excel.XlSortOrder.xlDescending, "Квартал");

            pivotTable.PivotFields("Месяц").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Месяц").Position = 3;
            pivotTable.PivotFields("Месяц").AutoSort(Excel.XlSortOrder.xlDescending, "Месяц");

            pivotTable.PivotFields("Неделя").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Неделя").Position = 4;
            pivotTable.PivotFields("Неделя").AutoSort(Excel.XlSortOrder.xlDescending, "Неделя");

            pivotTable.AddDataField(pivotTable.PivotFields("Номер ЛН"),
                "Кол-во Номер ЛН", Excel.XlConsolidationFunction.xlCount);

            pivotTable.PivotFields("Адрес").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            pivotTable.PivotFields("Адрес").Position = 1;

            pivotTable.HasAutoFormat = false;

            int wsPivoteColumnUsed = wsPivote.UsedRange.Columns.Count;

            wsPivote.Columns["B:" + ExcelGeneral.ColumnIndexToColumnLetter(wsPivoteColumnUsed)].Select();
            xlApp.Selection.ColumnWidth = 10;
            wsPivote.Range["B2:" + ExcelGeneral.ColumnIndexToColumnLetter(wsPivoteColumnUsed) + "2"].Select();
            xlApp.Selection.HorizontalAlignment = Excel.Constants.xlGeneral;
            xlApp.Selection.VerticalAlignment = Excel.Constants.xlTop;
            xlApp.Selection.WrapText = true;

            pivotTable.PivotFields("Месяц").ShowDetail = false;

            pivotTable.TableStyle2 = "PivotStyleMedium6";

            pivotTable.DisplayFieldCaptions = false;
            wb.ShowPivotTableFieldList = false;

            wsPivote.Range["A1"].Select();
        }
    }
}
