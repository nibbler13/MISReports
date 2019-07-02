using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
    class RecordsFromInsuranceCompanies : ExcelGeneral {
        public static bool Process(string resultFile) {
            if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
                out Excel.Worksheet ws))
                return false;

            int usedRows = ws.UsedRange.Rows.Count;

            ws.Range["A2:Q2"].Select();
            xlApp.Selection.Copy();
            ws.Range["A3:Q" + usedRows].Select();
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
            string pivotTableName = @"RecordsFromInsuranceCompaniesPivotTable";
            Excel.Worksheet wsPivote = wb.Sheets["Сводная Таблица"];

            wsPivote.Activate();

            Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
            Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

            pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

            pivotTable.PivotFields("Название страховой").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Название страховой").Position = 1;

            pivotTable.PivotFields("Имя оператора").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Имя оператора").Position = 2;

            pivotTable.PivotFields("Пациент").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Пациент").Position = 3;

            pivotTable.PivotFields("Отделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Отделение").Position = 4;

            pivotTable.PivotFields("Доктор").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Доктор").Position = 5;

            pivotTable.PivotFields("Дата назначения").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            pivotTable.PivotFields("Дата назначения").Position = 1;

            pivotTable.PivotFields("Название страховой").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
            pivotTable.PivotFields("Имя оператора").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
            pivotTable.PivotFields("Пациент").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
            pivotTable.PivotFields("Отделение").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
            pivotTable.PivotFields("Доктор").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
            pivotTable.PivotFields("Дата назначения").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };

            pivotTable.PivotFields("Название страховой").LayoutForm = Excel.XlLayoutFormType.xlTabular;
            pivotTable.PivotFields("Имя оператора").LayoutForm = Excel.XlLayoutFormType.xlTabular;
            pivotTable.PivotFields("Пациент").LayoutForm = Excel.XlLayoutFormType.xlTabular;
            pivotTable.PivotFields("Отделение").LayoutForm = Excel.XlLayoutFormType.xlTabular;
            pivotTable.PivotFields("Доктор").LayoutForm = Excel.XlLayoutFormType.xlTabular;
            pivotTable.PivotFields("Дата назначения").LayoutForm = Excel.XlLayoutFormType.xlTabular;

            pivotTable.AddDataField(
                pivotTable.PivotFields("SCHEDID"),
                "Кол-во",
                Excel.XlConsolidationFunction.xlCount);

            wb.ShowPivotTableFieldList = false;

            wsPivote.Range["A1"].Select();
        }
    }
}
