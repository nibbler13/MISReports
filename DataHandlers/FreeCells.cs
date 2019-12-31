using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class FreeCells : ExcelGeneral {

		//============================ FreeCells ============================
		public static bool Process(string resultFile, DateTime dateBeginOriginal, DateTime dateEnd) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws))
				return false;

			try {
				ws.Columns["C:C"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ";
				ws.Columns["C:C"].EntireColumn.AutoFit();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				AddPivotTableFreeCells(wb, ws, xlApp, false, dateBeginOriginal);
				//wb.Sheets["Данные"].Activate();
				//AddPivotTableFreeCells(wb, ws, xlApp, true, dateBeginOriginal, dateEnd);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			wb.Sheets["Сводная таблица"].Activate();
			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static void AddPivotTableFreeCells(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp,
			bool isMonth, DateTime date, DateTime? dateMonthEnd = null) {
			ws.Cells[1, 1].Select();

			string sheetName;
			if (isMonth) sheetName = "Сводная таблица текущий месяц";
			else sheetName = "Сводная таблица";

			string pivotTableName = @"PivotTable";
			Excel.Worksheet wsPivote = wb.Sheets[sheetName];
			wsPivote.Activate();

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 1;

			pivotTable.PivotFields("Пересечение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Пересечение").Position = 2;

			pivotTable.PivotFields("Отделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Отделение").Position = 3;

			pivotTable.PivotFields("Врач").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Врач").Position = 4;

			pivotTable.PivotFields("Должность").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Должность").Position = 5;

			pivotTable.AddDataField(pivotTable.PivotFields("Всего"), "(Всего)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.AddDataField(pivotTable.PivotFields("Занято"), "(Занято)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.AddDataField(pivotTable.PivotFields("% занятых слотов"), "(% занятых слотов)", Excel.XlConsolidationFunction.xlAverage);

			if (isMonth) {
				CultureInfo cultureInfoOriginal = Thread.CurrentThread.CurrentCulture;
				Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
				for (DateTime dateToHide = date; dateToHide.Date <= dateMonthEnd.Value.Date; dateToHide = dateToHide.AddDays(1)) {
					string pivotItem = dateToHide.ToShortDateString();
					Console.WriteLine("pivotItem: " + pivotItem);
					pivotTable.PivotFields("Дата").PivotItems(pivotItem).Visible = false;
				}
				Thread.CurrentThread.CurrentCulture = cultureInfoOriginal;
			} else {
				pivotTable.PivotFields("Дата").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
				pivotTable.PivotFields("Дата").Position = 1;
				pivotTable.PivotFields("Дата").AutoGroup();
				pivotTable.PivotFields("Дата").PivotFilters.Add2(Excel.XlPivotFilterType.xlAfter, null,
					date.AddDays(-1).ToShortDateString(), null, null, null, null, null, true);
				try { pivotTable.PivotFields("Месяцы").Orientation = Excel.XlPivotFieldOrientation.xlHidden; } catch (Exception) { }
			}

			pivotTable.RowGrand = false;
			pivotTable.ColumnGrand = false;
			pivotTable.DisplayFieldCaptions = false;

			pivotTable.PivotFields("(Занято)").NumberFormat = "0,00";
			pivotTable.PivotFields("(% занятых слотов)").NumberFormat = "0,0%";
			pivotTable.PivotSelect("'(% занятых слотов)'", Excel.XlPTSelectionMode.xlDataAndLabel, true);

			xlApp.Selection.FormatConditions.AddColorScale(3);
			xlApp.Selection.FormatConditions(xlApp.Selection.FormatConditions.Count).SetFirstPriority();

			xlApp.Selection.FormatConditions[1].ColorScaleCriteria[1].Type =
				Excel.XlConditionValueTypes.xlConditionValueLowestValue;
			xlApp.Selection.FormatConditions[1].ColorScaleCriteria[1].FormatColor.Color = 5287936;
			xlApp.Selection.FormatConditions[1].ColorScaleCriteria[1].FormatColor.TintAndShade = 0;


			xlApp.Selection.FormatConditions[1].ColorScaleCriteria[2].Type =
				Excel.XlConditionValueTypes.xlConditionValuePercentile;
			xlApp.Selection.FormatConditions[1].ColorScaleCriteria[2].Value = 65;
			xlApp.Selection.FormatConditions[1].ColorScaleCriteria[2].FormatColor.Color = 8711167;
			xlApp.Selection.FormatConditions[1].ColorScaleCriteria[2].FormatColor.TintAndShade = 0;

			xlApp.Selection.FormatConditions[1].ColorScaleCriteria[3].Type =
				Excel.XlConditionValueTypes.xlConditionValueHighestValue;
			xlApp.Selection.FormatConditions[1].ColorScaleCriteria[3].FormatColor.Color = 255;
			xlApp.Selection.FormatConditions[1].ColorScaleCriteria[3].FormatColor.TintAndShade = 0;

			xlApp.Selection.FormatConditions[1].ScopeType = Excel.XlPivotConditionScope.xlDataFieldScope;

			pivotTable.PivotFields("Порядок сортировки").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Порядок сортировки").Position = 1;
			pivotTable.PivotFields("Порядок сортировки").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Порядок сортировки").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Отделение").ShowDetail = false;
			pivotTable.PivotFields("Пересечение").ShowDetail = false;
			pivotTable.PivotFields("Филиал").ShowDetail = false;

			wsPivote.Range["A1"].Select();
			wb.ShowPivotTableFieldList = false;
		}

	}
}
