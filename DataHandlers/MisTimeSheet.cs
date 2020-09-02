using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class MisTimeSheet : ExcelGeneral {
        public static bool Process(string resultFile) {
            if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
                out Excel.Worksheet ws))
                return false;

            int usedRows = ws.UsedRange.Rows.Count;

            ws.Range["A2:P2"].Select();
            xlApp.Selection.Copy();
            ws.Range["A3:P" + usedRows].Select();
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
            string pivotTableName = @"MisTimeSheetPivotTable";
            Excel.Worksheet wsPivote = wb.Sheets["Сводная Таблица"];

            wsPivote.Activate();

            Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
            Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

            pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

            pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Филиал").Position = 1;

            pivotTable.PivotFields("Полное имя доктора").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Полное имя доктора").Position = 2;

            pivotTable.PivotFields("Идентификатор доктора").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Идентификатор доктора").Position = 3;

            pivotTable.PivotFields("Должность").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Должность").Position = 4;

            pivotTable.PivotFields("Должность (справочник)").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Должность (справочник)").Position = 5;

            pivotTable.PivotFields("Дата графика работ").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            pivotTable.PivotFields("Дата графика работ").Position = 1;

			pivotTable.AddDataField(pivotTable.PivotFields("Кол-во часов (план)"), "Сумма кол-во часов (план)", Excel.XlConsolidationFunction.xlSum);
            pivotTable.PivotFields("Сумма кол-во часов (план)").NumberFormat = "# ##0,00";

			pivotTable.PivotFields("Дата графика работ").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Должность (справочник)").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Должность").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Идентификатор доктора").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Полное имя доктора").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Филиал").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };

            pivotTable.PivotFields("Дата графика работ").LayoutForm = Excel.XlLayoutFormType.xlTabular;
			pivotTable.PivotFields("Должность (справочник)").LayoutForm = Excel.XlLayoutFormType.xlTabular;
			pivotTable.PivotFields("Должность").LayoutForm = Excel.XlLayoutFormType.xlTabular;
			pivotTable.PivotFields("Идентификатор доктора").LayoutForm = Excel.XlLayoutFormType.xlTabular;
			pivotTable.PivotFields("Полное имя доктора").LayoutForm = Excel.XlLayoutFormType.xlTabular;
			pivotTable.PivotFields("Филиал").LayoutForm = Excel.XlLayoutFormType.xlTabular;

            //pivotTable.PivotFields("Группа филиалов").ShowDetail = false;
            //pivotTable.PivotFields("Название").ShowDetail = false;

            wb.ShowPivotTableFieldList = false;

            wsPivote.Range["A1"].Select();
        }
    }
}
