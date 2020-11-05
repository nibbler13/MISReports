using MISReports.ExcelHandlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.DataHandlers {
	class RecordsCountFrontOffice : ExcelGeneral {
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

			int rowsUsed = ws.UsedRange.Rows.Count;

			wsPivote.Activate();

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, "Данные!R1C1:R" + rowsUsed + "C4", 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			pivotTable.PivotFields("ФИО сотрудника").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ФИО сотрудника").Position = 1;

			pivotTable.PivotFields("Должность").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Должность").Position = 2;

			pivotTable.PivotFields("Дата").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
			pivotTable.PivotFields("Дата").Position = 1;

			pivotTable.AddDataField(pivotTable.PivotFields("Кол-во клиентов"), "Кол-во клиентов ", Excel.XlConsolidationFunction.xlSum);

			pivotTable.PivotFields("ФИО сотрудника").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("ФИО сотрудника").LayoutForm = Excel.XlLayoutFormType.xlTabular;
			pivotTable.PivotFields("Дата").NumberFormat = "[$-ru-RU]Д МММ;@";

			wb.ShowPivotTableFieldList = false;
			pivotTable.DisplayFieldCaptions = false;

			wsPivote.Range["A1"].Select();
		}
	}
}
