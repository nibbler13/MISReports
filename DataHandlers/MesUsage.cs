using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class MesUsage : ExcelGeneral {

		//============================ MesUsage ============================
		public static bool Process(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws))
				return false;

			try {
				ws.Activate();
				ws.Columns["C:C"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ";
				ws.Columns["P:P"].Select();
				xlApp.Selection.NumberFormat = "0%";
				ws.Range["A1"].Select();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				MesUsageAddPivotTable(wb, ws, xlApp);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static void MesUsageAddPivotTable(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			ws.Cells[1, 1].Select();

			string pivotTableName = @"MesUsagePivotTable";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная таблица"];

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			pivotTable.PivotFields("Тип приема").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
			pivotTable.PivotFields("Тип приема").Position = 1;

			pivotTable.PivotFields("Тип оплаты приема").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
			pivotTable.PivotFields("Тип оплаты приема").Position = 2;

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 1;

			pivotTable.PivotFields("Подразделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Подразделение").Position = 2;

			pivotTable.PivotFields("ФИО врача").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ФИО врача").Position = 3;

			pivotTable.AddDataField(pivotTable.PivotFields("Прием"),
				"Кол-во приемов, для которых загружен список МЭС", Excel.XlConsolidationFunction.xlSum);
			pivotTable.AddDataField(pivotTable.PivotFields("Есть направление, созданное с использованием МЭС"),
				"Кол-во приемов с направлением, созданным с использованием МЭС", Excel.XlConsolidationFunction.xlSum);

			pivotTable.CalculatedFields().Add("% приемов с направлением, созданным с использованием МЭС",
				"='Есть направление, созданное с использованием МЭС' /Прием", true);
			pivotTable.PivotFields("% приемов с направлением, созданным с использованием МЭС").Orientation =
				Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю % приемов с направлением, созданным с использованием МЭС").Caption =
				" % приемов с направлением, созданным с использованием МЭС";
			pivotTable.PivotFields(" % приемов с направлением, созданным с использованием МЭС").NumberFormat = "0,00%";

			pivotTable.AddDataField(pivotTable.PivotFields("Есть направление, созданное самостоятельно"),
				"Кол-во приемов с направлениями, созданными самостоятельно", Excel.XlConsolidationFunction.xlSum);

			pivotTable.CalculatedFields().Add("% приемов с направлениями, соответствующими МЭС, но созданных самостоятельно",
				"='Есть направление, созданное самостоятельно' /Прием", true);
			pivotTable.PivotFields("% приемов с направлениями, соответствующими МЭС, но созданных самостоятельно").Orientation =
				Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю % приемов с направлениями, соответствующими МЭС, но созданных самостоятельно").Caption =
				" % приемов с направлениями, соответствующими МЭС, но созданных самостоятельно";
			pivotTable.PivotFields(" % приемов с направлениями, соответствующими МЭС, но созданных самостоятельно").NumberFormat = "0,00%";

			pivotTable.AddDataField(pivotTable.PivotFields("Услуги из всех направлений соответсвуют МЭС на 100%"),
				"Кол-во приемов, обязательные услуги МЭС соответствуют в направлениях на 100%", Excel.XlConsolidationFunction.xlSum);

			pivotTable.CalculatedFields().Add("% приемов, обязательные услуги МЭС в направлениях соответствуют на 100%",
				"='Услуги из всех направлений соответсвуют МЭС на 100%' /Прием", true);
			pivotTable.PivotFields("% приемов, обязательные услуги МЭС в направлениях соответствуют на 100%").Orientation =
				Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю % приемов, обязательные услуги МЭС в направлениях соответствуют на 100%").Caption =
				" % приемов, обязательные услуги МЭС в направлениях соответствуют на 100%";
			pivotTable.PivotFields(" % приемов, обязательные услуги МЭС в направлениях соответствуют на 100%").NumberFormat = "0,00%";

			pivotTable.AddDataField(pivotTable.PivotFields("% Соответствия МЭС"),
				"Средний % соответствия обязательных услуг МЭС услугам в направлениях", Excel.XlConsolidationFunction.xlAverage);
			pivotTable.PivotFields("Средний % соответствия обязательных услуг МЭС услугам в направлениях").NumberFormat = "0,00%";

			wsPivote.Activate();
			wsPivote.Columns["B:I"].Select();
			xlApp.Selection.ColumnWidth = 20;
			wsPivote.Range["B4:I4"].Select();
			xlApp.Selection.VerticalAlignment = Excel.Constants.xlTop;
			xlApp.Selection.WrapText = true;

			pivotTable.PivotFields("ФИО врача").AutoSort(Excel.XlSortOrder.xlDescending,
				"Средний % соответствия обязательных услуг МЭС услугам в направлениях");
			pivotTable.PivotFields("Подразделение").AutoSort(Excel.XlSortOrder.xlDescending,
				"Средний % соответствия обязательных услуг МЭС услугам в направлениях");
			pivotTable.PivotFields("Филиал").AutoSort(Excel.XlSortOrder.xlDescending,
				"Средний % соответствия обязательных услуг МЭС услугам в направлениях");

			int rowCount = wsPivote.UsedRange.Rows.Count;
			AddInteriorColor(wsPivote.Range["C4:D" + rowCount], Excel.XlThemeColor.xlThemeColorAccent4);
			AddInteriorColor(wsPivote.Range["E4:F" + rowCount], Excel.XlThemeColor.xlThemeColorAccent5);
			AddInteriorColor(wsPivote.Range["G4:H" + rowCount], Excel.XlThemeColor.xlThemeColorAccent6);

			wsPivote.Range["A1"].Select();

			pivotTable.HasAutoFormat = false;

			pivotTable.PivotFields("Подразделение").ShowDetail = false;
			pivotTable.PivotFields("Филиал").ShowDetail = false;

			wb.ShowPivotTableFieldList = false;
		}

	}
}
