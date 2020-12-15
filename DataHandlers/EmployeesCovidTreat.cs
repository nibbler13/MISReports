using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class EmployeesCovidTreat : ExcelGeneral {
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

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}


		private static void AddPivotTable(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			string pivotTableName = @"PivotTableEmployeesCovidTreat";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная таблица"];

			int rowsUsed = ws.UsedRange.Rows.Count;

			wsPivote.Activate();

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, "Данные!R1C1:R" + rowsUsed + "C14", 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 1;

			pivotTable.PivotFields("Диагноз МКБ").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Диагноз МКБ").Position = 2;

			pivotTable.PivotFields("Диагноз").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Диагноз").Position = 3;

			pivotTable.PivotFields("ФИО пациента").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ФИО пациента").Position = 4;

			pivotTable.PivotFields("Отделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Отделение").Position = 5;

			pivotTable.PivotFields("ФИО пациента").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("ФИО пациента").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Диагноз").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Диагноз").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Диагноз МКБ").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Диагноз МКБ").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			//pivotTable.PivotFields("Запись, Должность").ShowDetail = false;
			//pivotTable.PivotFields("Филиал").ShowDetail = false;
			//pivotTable.PivotFields("ФИО пользователя").ShowDetail = false;

			wb.ShowPivotTableFieldList = false;
			pivotTable.DisplayFieldCaptions = false;

			wsPivote.Columns["B:B"].ColumnWidth = 60;

			wsPivote.Range["A1"].Select();
		}
	}
}
