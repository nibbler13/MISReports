using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class NonAppearance : ExcelGeneral {

		//============================ NonAppearance ============================
		public static bool Process(string resultFile, DataTable dataTable) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws))
				return false;

			try {
				int usedRows = ws.UsedRange.Rows.Count;
				ws.Columns["B:B"].NumberFormat = "ДД.ММ.ГГГГ";
				ws.Range["A1"].Select();
				ws.Columns["K:K"].NumberFormat = "0,00%";
				ws.Columns["M:M"].NumberFormat = "0,00%";
				ws.Range["N2"].Select();
				xlApp.ActiveCell.FormulaR1C1 = "=RC[-4]+RC[-2]";
				xlApp.Selection.AutoFill(ws.Range["N2:N" + usedRows]);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				NonAppearanceAddPivotTablePatientsWithProblem(wb, ws, xlApp);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				NonAppearanceAddStatistics(wb, xlApp, dataTable);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				NonAppearanceAddPivotTableGeneral(wb, ws, xlApp);

				ws = wb.Sheets["Сводная таблица"];
				ws.Activate();
				ws.Columns["B:G"].ColumnWidth = 15;
				ws.Range["B1:G1"].Select();
				xlApp.Selection.WrapText = true;
				ws.Range["A1"].Select();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static void NonAppearanceAddStatistics(Excel.Workbook wb, Excel.Application xlApp, DataTable dataTable) {
			SortedDictionary<string, SortedDictionary<string, ItemNonAppearanceStatistic>> dict =
				new SortedDictionary<string, SortedDictionary<string, ItemNonAppearanceStatistic>> {
					{ "Всего", new SortedDictionary<string, ItemNonAppearanceStatistic>() }
				};

			foreach (DataRow row in dataTable.Rows) {
				try {
					string filial = row["FILIAL_SHORTNAME"].ToString();
					string recordType = row["ИСТОЧНИК ЗАПИСИ"].ToString();
					int recordsCount = Convert.ToInt32(row["PATIENTS_TOTAL"].ToString());
					int nonAppearanceCount = Convert.ToInt32(row["MARKS_WITHOUT_TREATMENTS"].ToString()) +
						Convert.ToInt32(row["WOUT_MARKS_WITHOUT_TREATMENTS"].ToString());

					if (!dict.ContainsKey(filial))
						dict.Add(filial, new SortedDictionary<string, ItemNonAppearanceStatistic>());

					foreach (string key in new string[] { filial, "Всего" }) {
						if (!dict[key].ContainsKey(recordType))
							dict[key].Add(recordType, new ItemNonAppearanceStatistic(recordType));

						dict[key][recordType].AddValues(recordsCount, nonAppearanceCount);
					}
				} catch (Exception e) {
					Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
				}
			}

			Excel.Worksheet ws = wb.Sheets["Статистика"];
			ws.Activate();
			int currentRow = 2;
			int colorIndex = 20;
			int previousBlockRow = currentRow;

			foreach (KeyValuePair<string, SortedDictionary<string, ItemNonAppearanceStatistic>> keyValuePair in dict) {
				foreach (KeyValuePair<string, ItemNonAppearanceStatistic> innerKeyValuePair in keyValuePair.Value) {
					ws.Cells[currentRow, 1].Value2 = keyValuePair.Key;
					ws.Cells[currentRow, 2].Value2 = innerKeyValuePair.Key;
					ws.Cells[currentRow, 3].Value2 = innerKeyValuePair.Value.NonAppearanceCount;
					ws.Cells[currentRow, 4].Value2 = innerKeyValuePair.Value.RecordsCount;
					ws.Cells[currentRow, 5].Value2 = (double)innerKeyValuePair.Value.NonAppearanceCount / (double)innerKeyValuePair.Value.RecordsCount;
					currentRow++;
				}

				ws.Range["A" + previousBlockRow + ":E" + (currentRow - 1)].Select();
				foreach (Excel.XlBordersIndex border in new Excel.XlBordersIndex[] {
					Excel.XlBordersIndex.xlInsideHorizontal,
					Excel.XlBordersIndex.xlInsideVertical}) {
					xlApp.Selection.Borders[border].LineStyle = Excel.XlLineStyle.xlDot;
					xlApp.Selection.Borders[border].ColorIndex = 0;
					xlApp.Selection.Borders[border].TintAndShade = 0;
					xlApp.Selection.Borders[border].Weight = Excel.XlBorderWeight.xlThin;
				}

				foreach (Excel.XlBordersIndex border in new Excel.XlBordersIndex[] {
					Excel.XlBordersIndex.xlEdgeBottom,
					Excel.XlBordersIndex.xlEdgeLeft,
					Excel.XlBordersIndex.xlEdgeRight,
					Excel.XlBordersIndex.xlEdgeTop}) {
					xlApp.Selection.Borders[border].LineStyle = Excel.XlLineStyle.xlDouble;
					xlApp.Selection.Borders[border].ColorIndex = 0;
					xlApp.Selection.Borders[border].TintAndShade = 0;
					xlApp.Selection.Borders[border].Weight = Excel.XlBorderWeight.xlThin;
				}

				xlApp.Selection.Interior.ColorIndex = colorIndex;
				previousBlockRow = currentRow;
				colorIndex = colorIndex == 19 ? 20 : 19;
			}

			ws.Cells[1, 1].Select();
			wb.Sheets["Данные"].Activate();
		}

		private class ItemNonAppearanceStatistic {
			public string Name { get; private set; }
			public int RecordsCount { get; private set; }
			public int NonAppearanceCount { get; private set; }

			public ItemNonAppearanceStatistic(string name) {
				Name = name;
				RecordsCount = 0;
				NonAppearanceCount = 0;
			}

			public void AddValues(int recordsCount, int nonAppearanceCount) {
				RecordsCount += recordsCount;
				NonAppearanceCount += nonAppearanceCount;
			}
		}

		private static void NonAppearanceAddPivotTableGeneral(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			ws.Cells[1, 1].Select();

			string pivotTableName = @"NonAppearancePivotTable";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная Таблица"];

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 1;

			pivotTable.PivotFields("Подразделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Подразделение").Position = 2;

			pivotTable.PivotFields("ФИО доктора").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ФИО доктора").Position = 3;

			pivotTable.PivotFields("Дата лечения").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Дата лечения").Position = 4;

			pivotTable.AddDataField(pivotTable.PivotFields("Записано пациентов"),
				"Всего записано пациентов", Excel.XlConsolidationFunction.xlSum);

			pivotTable.AddDataField(pivotTable.PivotFields("Отметки без лечений"),
				"Отметки без лечения (регистратура +, врач – )", Excel.XlConsolidationFunction.xlSum);
			pivotTable.CalculatedFields().Add("Общий % Неявок - Отметки без лечений",
				"= 'Отметки без лечений'/'Записано пациентов'", true);
			pivotTable.PivotFields("Общий % Неявок - Отметки без лечений").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю Общий % Неявок - Отметки без лечений").NumberFormat = "0,00%";
			pivotTable.PivotFields("Сумма по полю Общий % Неявок - Отметки без лечений").Caption =
				"% Неявок - Отметки без лечений (регистратура +, врач – )";

			pivotTable.AddDataField(pivotTable.PivotFields("Без отметок и без лечений"),
				"Без отметок и лечения (регистратура -, врач -)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.CalculatedFields().Add("Общий % Неявок - Без отметок и без лечений",
				"= 'Без отметок и без лечений'/'Записано пациентов'", true);
			pivotTable.PivotFields("Общий % Неявок - Без отметок и без лечений").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю Общий % Неявок - Без отметок и без лечений").NumberFormat = "0,00%";
			pivotTable.PivotFields("Сумма по полю Общий % Неявок - Без отметок и без лечений").Caption =
				"% Неявок - Без отметок и без лечений (регистратура -, врач -)";

			pivotTable.CalculatedFields().Add("Общий % Неявки",
				"= ('Отметки без лечений' +'Без отметок и без лечений' )/'Записано пациентов'", true);
			pivotTable.PivotFields("Общий % Неявки").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю Общий % Неявки").NumberFormat = "0,00%";
			pivotTable.PivotFields("Сумма по полю Общий % Неявки").Caption = "% Неявки";

			pivotTable.HasAutoFormat = false;

			pivotTable.PivotFields("ФИО доктора").ShowDetail = false;
			pivotTable.PivotFields("Подразделение").ShowDetail = false;
			pivotTable.PivotFields("Филиал").ShowDetail = false;

			pivotTable.DisplayFieldCaptions = false;
			wb.ShowPivotTableFieldList = false;
		}

		private static void NonAppearanceAddPivotTablePatientsWithProblem(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			ws.Cells[1, 1].Select();

			string pivotTableName = @"PatientsWithProblem";
			Excel.Worksheet wsPivote = wb.Sheets["Пациенты с неявками"];

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 1;

			pivotTable.PivotFields("Подразделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Подразделение").Position = 2;

			pivotTable.PivotFields("ФИО доктора").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ФИО доктора").Position = 3;

			pivotTable.PivotFields("Дата лечения").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Дата лечения").Position = 4;
			pivotTable.PivotFields("Дата лечения").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Дата лечения").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("ФИО пациента").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ФИО пациента").Position = 5;
			pivotTable.PivotFields("ФИО пациента").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("ФИО пациента").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("История болезни пациента").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("История болезни пациента").Position = 6;
			pivotTable.PivotFields("История болезни пациента").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("История болезни пациента").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Номер телефона пациента").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Номер телефона пациента").Position = 7;
			pivotTable.PivotFields("Номер телефона пациента").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Номер телефона пациента").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Отметки без лечений + Без отметок и без лечений").Orientation =
				Excel.XlPivotFieldOrientation.xlPageField;
			pivotTable.PivotFields("Отметки без лечений + Без отметок и без лечений").Position = 1;

			pivotTable.PivotFields("Отметки без лечений + Без отметок и без лечений").CurrentPage = "(ALL)";
			pivotTable.PivotFields("Отметки без лечений + Без отметок и без лечений").PivotItems("0").Visible = false;
			pivotTable.PivotFields("Отметки без лечений + Без отметок и без лечений").EnableMultiplePageItems = true;

			pivotTable.HasAutoFormat = false;

			//pivotTable.PivotFields("ФИО доктора").ShowDetail = false;
			//pivotTable.PivotFields("Подразделение").ShowDetail = false;
			pivotTable.PivotFields("Филиал").ShowDetail = false;

			//pivotTable.DisplayFieldCaptions = false;
			wb.ShowPivotTableFieldList = false;
		}

	}
}
