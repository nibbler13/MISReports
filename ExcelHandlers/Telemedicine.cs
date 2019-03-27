using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class Telemedicine : ExcelGeneral {


		//============================ Telemedicine ============================
		public static bool Process(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws))
				return false;

			try {
				ws.Columns["C:C"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ";
				ws.Columns["I:I"].ColumnWidth = 10;
				ws.Columns["I:I"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ";
				ws.Columns["I:I"].ColumnWidth = 10;
				ws.Range["A1"].Select();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				TelemedicineAddPivotTable(wb, ws, xlApp);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			wb.Sheets["Сводная таблица"].Activate();
			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static void TelemedicineAddPivotTable(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			ws.Cells[1, 1].Select();

			string pivotTableName = @"TelemedicinePivotTable";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная Таблица"];

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			pivotTable.PivotFields("FILIAL_SHORTNAME").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("FILIAL_SHORTNAME").Position = 1;

			pivotTable.PivotFields("SERVICE_TYPE").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("SERVICE_TYPE").Position = 2;

			pivotTable.PivotFields("CLIENT_CATEGORY").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
			pivotTable.PivotFields("CLIENT_CATEGORY").Position = 1;

			pivotTable.AddDataField(pivotTable.PivotFields("CLIENT_HITSNUM"), "Кол-во", Excel.XlConsolidationFunction.xlCount);
			pivotTable.DisplayFieldCaptions = false;
			wb.ShowPivotTableFieldList = false;
			pivotTable.ShowDrillIndicators = false;


			//ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
			//	"Данные!R1C1:R38C16", Version:=6).CreatePivotTable TableDestination:= _
			//	"Сводная таблица!R1C1", TableName:="Сводная таблица1", DefaultVersion:=6
			//Sheets("Сводная таблица").Select
			//Cells(1, 1).Select
			//With ActiveSheet.PivotTables("Сводная таблица1").PivotFields("SERVICE_TYPE")
			//	.Orientation = xlRowField
			//	.Position = 1
			//End With
			//With ActiveSheet.PivotTables("Сводная таблица1").PivotFields("CLIENT_CATEGORY")
			//	.Orientation = xlColumnField
			//	.Position = 1
			//End With
			//ActiveSheet.PivotTables("Сводная таблица1").AddDataField ActiveSheet. _
			//	PivotTables("Сводная таблица1").PivotFields("CLIENT_HITSNUM"), _
			//	"Сумма по полю CLIENT_HITSNUM", xlSum
			//With ActiveSheet.PivotTables("Сводная таблица1").PivotFields( _
			//	"Сумма по полю CLIENT_HITSNUM")
			//	.Caption = "Количество по полю CLIENT_HITSNUM"
			//	.Function = xlCount
			//End With
			//ActiveSheet.PivotTables("Сводная таблица1").DisplayFieldCaptions = False
			//ActiveWorkbook.ShowPivotTableFieldList = False
			//ActiveSheet.PivotTables("Сводная таблица1").ShowDrillIndicators = False
		}

	}
}
