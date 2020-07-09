using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
    class Promo : ExcelGeneral {
        public static bool Process(string resultFile) {
            if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
                out Excel.Worksheet ws))
                return false;

            int usedRows = ws.UsedRange.Rows.Count;

            ws.Range["A2:F2"].Select();
            xlApp.Selection.Copy();
            ws.Range["A3:F" + usedRows].Select();
            xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
            ws.Range["A1"].Select();

            try {
                AddPivotTable(wb, ws, xlApp);
            } catch (Exception e) {
                Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
            }

            SaveAndCloseWorkbook(xlApp, wb, ws);

            return true;
        }

        private static void AddPivotTable(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
            string pivotTableName = @"PromoPivotTable";
            Excel.Worksheet wsPivote = wb.Sheets["Сводная Таблица"];

            wsPivote.Activate();

            Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
            Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

            pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

            pivotTable.PivotFields("Название").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Название").Position = 1;

            pivotTable.PivotFields("Группа филиалов").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Группа филиалов").Position = 2;

            pivotTable.PivotFields("Услуга").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Услуга").Position = 3;

            //pivotTable.PivotFields("Название").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
            //pivotTable.PivotFields("Группа филиалов").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
            //pivotTable.PivotFields("Услуга").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };

            //pivotTable.PivotFields("Название").LayoutForm = Excel.XlLayoutFormType.xlTabular;
            //pivotTable.PivotFields("Группа филиалов").LayoutForm = Excel.XlLayoutFormType.xlTabular;
            //pivotTable.PivotFields("Услуга").LayoutForm = Excel.XlLayoutFormType.xlTabular;

            pivotTable.PivotFields("Группа филиалов").ShowDetail = false;
            pivotTable.PivotFields("Название").ShowDetail = false;

            wb.ShowPivotTableFieldList = false;

            wsPivote.Range["A1"].Select();
        }
    }
}
