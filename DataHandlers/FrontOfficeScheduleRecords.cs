using MISReports.ExcelHandlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.DataHandlers {
	class FrontOfficeScheduleRecords : ExcelGeneral {
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
			string pivotTableName = @"PivotTableScheduleRecords";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная таблица"];

			int rowsUsed = ws.UsedRange.Rows.Count;

			wsPivote.Activate();

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, "Данные!R1C1:R" + rowsUsed + "C16", 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			pivotTable.PivotFields("ФИО пользователя").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ФИО пользователя").Position = 1;

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 2;

			pivotTable.AddDataField(pivotTable.PivotFields("Дата и время записи"), "Кол-во записей", Excel.XlConsolidationFunction.xlCount);
			pivotTable.PivotFields("Кол-во записей").NumberFormat = "# ##0";

			pivotTable.AddDataField(pivotTable.PivotFields("По направлению?"), "По направлению", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("По направлению").NumberFormat = "# ##0";

			pivotTable.AddDataField(pivotTable.PivotFields("Прием состоялся?"), "Прием состоялся", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Прием состоялся").NumberFormat = "# ##0";

			pivotTable.AddDataField(pivotTable.PivotFields("Сумма, всего"), "Сумма оказанных услуг (руб)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Сумма оказанных услуг (руб)").NumberFormat = "# ##0,00 ?";

			//pivotTable.PivotFields("Запись, Должность").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			//pivotTable.PivotFields("Запись, Должность").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			//pivotTable.PivotFields("Филиал").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			//pivotTable.PivotFields("Филиал").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			//pivotTable.PivotFields("Дата").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			//pivotTable.PivotFields("Дата").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			//pivotTable.PivotFields("Запись, Должность").ShowDetail = false;
			//pivotTable.PivotFields("Филиал").ShowDetail = false;
			pivotTable.PivotFields("ФИО пользователя").ShowDetail = false;

			wb.ShowPivotTableFieldList = false;
			//pivotTable.DisplayFieldCaptions = false;

			wsPivote.Range["A1"].Select();
		}
	}
}
