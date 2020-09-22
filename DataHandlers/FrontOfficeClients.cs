using MISReports.ExcelHandlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.DataHandlers {
	class FrontOfficeClients : ExcelGeneral {
		public static bool Process(string resultFile) {
			if (!CopyFormatting(resultFile))
				return false;

			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws))
				return false;

			try {
				AddPivotTable(wb, ws, xlApp);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			wb.Sheets["Сводная таблица"].Activate();
			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

        private static void AddPivotTable(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
            string pivotTableName = @"FrontOfficeClientsPivotTable";
            Excel.Worksheet wsPivote = wb.Sheets["Сводная Таблица"];

            wsPivote.Activate();

            Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
            Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

            pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

            pivotTable.PivotFields("Дата назначения").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Дата назначения").Position = 1;

            pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Филиал").Position = 2;

            pivotTable.PivotFields("Отделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("Отделение").Position = 3;

            pivotTable.AddDataField(pivotTable.PivotFields("№ ИБ"), "Кол-во записей", Excel.XlConsolidationFunction.xlCount);

            pivotTable.PivotFields("Филиал").ShowDetail = false;

            wb.ShowPivotTableFieldList = false;
            pivotTable.DisplayFieldCaptions = false;

            wsPivote.Range["A1"].Select();
        }
    }
}
