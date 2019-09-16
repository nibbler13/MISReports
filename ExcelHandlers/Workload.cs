using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class Workload : ExcelGeneral {
		public static bool Process(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws, "Услуги Мет. 1"))
				return false;

			try {
				ws.Activate();
				ws.Range["CC2:CE2"].Select();
				xlApp.Selection.AutoFill(ws.Range["CC2:CE" + ws.UsedRange.Rows.Count]);
				ws.Range["CC3:CE3"].Select();
				xlApp.Selection.AutoFill(ws.Range["CC2:CE3"]);
				ws.Range["A2:CB2"].Select();
				xlApp.Selection.Copy();
				ws.Range["A3:CB" + ws.UsedRange.Rows.Count].Select();
				xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
				ws.Range["A1"].Select();

				ws = wb.Sheets["Искл. услуги"];
				ws.Activate();
				ws.Range["A2:K2"].Select();
				xlApp.Selection.Copy();
				ws.Range["A3:K" + ws.UsedRange.Rows.Count].Select();
				xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
				ws.Range["A1"].Select();

				ws = wb.Sheets["Расчет"];
				ws.Activate();
				ws.Range["AA2:AM2"].Select();
				xlApp.Selection.AutoFill(ws.Range["AA2:AM" + ws.UsedRange.Rows.Count]);
				ws.Range["AA3:AM3"].Select();
				xlApp.Selection.AutoFill(ws.Range["AA2:AM3"]);
				ws.Range["A2:Z2"].Select();
				xlApp.Selection.Copy();
				ws.Range["A3:Z" + ws.UsedRange.Rows.Count].Select();
				xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
				ws.Range["Y2:Y2"].Select();
				xlApp.Selection.AutoFill(ws.Range["Y2:Y" + ws.UsedRange.Rows.Count]);

				List<string> deptsToExclude = new List<string> {
					"АНЕСТЕЗИОЛОГИЯ-РЕАНИМАТОЛОГИЯ",
					"ДЕЖУРНЫЙ ВРАЧ",
					"Дежурный врач детский",
					"ПРОЦЕДУРНЫЙ",
					"Процедурный кабинет детский",
					"ФИЗИОПРОЦЕДУРЫ",
					"ТЕЛЕМЕДИЦИНА",
					"Аппаратная коррекция зрения",
					"Процедурный (детский)",
					"Физиопроцедуры детские",
					"ЗДРАВПУНКТ",
					"ОБЩЕЕ",
					"ОМС",
					"ПРЕДРЕЙСОВЫЙ ОСМОТР",
					"СКОРАЯ МЕДИЦИНСКАЯ ПОМОЩЬ"
				};

				List<string> docPostsToExclude = new List<string>() {
					"Администратор",
					"Администратор (старший)",
					"Администратор-регистратор",
					"Администратор-регистратор-кассир",
					"Администрация",
					"Акушерка",
					"Архивариус",
					"Врач (общее)",
					"Главный врач",
					"Заведующий регистратурой",
					"Зам.главного врача по КЭР",
					"Лаборант",
					"Лаборатория",
					"Медицинская сестра",
					"Медицинская сестра (косметология)",
					"Медицинский брат",
					"Сотрудник КДЛ",
					"Сотрудник УК",
					"Фельдшер"
				};

				for (int row = 2; row < ws.UsedRange.Rows.Count; row++) {
					try {
						string department = ws.Range["F" + row].Value2;
						string docPost = ws.Range["K" + row].Value2;
						string filialCode = Convert.ToString(ws.Range["C" + row].Value2);

						if (department.ToLower().Equals("рефлексотерапия")) {
							double filID = ws.Range["C" + row].Value;
							if (filID == 1 || filID == 5 || filID == 12) {
								double chairsCount = 2;
								if (filID == 5) {
									chairsCount = 3;
									ws.Range["Y" + row].Value2 = "4";
									ws.Range["Y" + row].Interior.ColorIndex = 45;
								}

								double timeDS = ws.Range["L" + row].Value2;
								ws.Range["L" + row].FormulaLocal = "=" + timeDS + "/" + chairsCount;
								ws.Range["L" + row].Interior.ColorIndex = 45;

								double timeSchRez = ws.Range["N" + row].Value2;
								ws.Range["N" + row].FormulaLocal = "=" + timeSchRez + "/" + chairsCount;
								ws.Range["N" + row].Interior.ColorIndex = 45;
							}
						}

						if (string.IsNullOrEmpty(department))
							continue;

						if (deptsToExclude.Contains(department)) {
							ws.Range["AL" + row].Value2 = 1;
							continue;
						}

						if (string.IsNullOrEmpty(docPost))
							continue;

						if (docPostsToExclude.Contains(docPost)) {
							string deptLow = department.ToLower();
							if (!(deptLow.Contains("массаж") ||
								deptLow.Contains("водолечение") ||
								deptLow.Contains("грязелечение") ||
								deptLow.Contains("лечебная физкультура") ||
								deptLow.Contains("медицинская реабилитация"))) {
								ws.Range["AL" + row].Value2 = 1;
								continue;
							}
						}

						if (docPost.Equals("Рентгенолаборант") &&
							!filialCode.Equals("17") &&
							!filialCode.Equals("15")) {
							ws.Range["AL" + row].Value2 = 1;
							continue;
						}

						if (docPost.Equals("Мануальный терапевт") &&
							((string)ws.Range["H" + row].Value2).StartsWith("Пеньтковский"))
							ws.Range["AL" + row].Value2 = 1;

					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					}
				}

				ws.Columns["AM:AM"].Select();
				xlApp.Selection.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: "=ДЛСТР(СЖПРОБЕЛЫ(AM1))=0");
				xlApp.Selection.FormatConditions(xlApp.Selection.FormatConditions.Count).SetFirstPriority();
				xlApp.Selection.FormatConditions(1).Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
				xlApp.Selection.FormatConditions(1).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
				xlApp.Selection.FormatConditions(1).Interior.TintAndShade = 0;

				//xlApp.Selection.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, "=20");
				//xlApp.Selection.FormatConditions(2).Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
				//xlApp.Selection.FormatConditions(2).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
				//xlApp.Selection.FormatConditions(2).Interior.TintAndShade = 0.799981688894314;

				//xlApp.Selection.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, "=150");
				//xlApp.Selection.FormatConditions(3).Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
				//xlApp.Selection.FormatConditions(3).Interior.Color = 65535;
				//xlApp.Selection.FormatConditions(3).Interior.TintAndShade = 0;

				ws.Range["A1"].Select();

				xlApp.ActiveWindow.ScrollColumn = 8;
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				WorkloadAddPivotTable(wb, ws, xlApp);
			} catch (Exception e) {
				Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static void WorkloadAddPivotTable(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			ws.Cells[1, 1].Select();

			string pivotTableName = @"WorkloadPivotTable";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная таблица"];

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable.PivotFields("ОТДЕЛЕНИЕ").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ОТДЕЛЕНИЕ").Position = 1;

			pivotTable.PivotFields("ФИЛИАЛ").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
			pivotTable.PivotFields("ФИЛИАЛ").Position = 1;

			pivotTable.AddDataField(pivotTable.PivotFields("Загрузка"),
				"Средняя загрузка", Excel.XlConsolidationFunction.xlAverage);
			pivotTable.PivotFields("Средняя загрузка").NumberFormat = "# ##0,00";

			pivotTable.PivotFields("Не учитывать (=1)").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
			pivotTable.PivotFields("Не учитывать (=1)").Position = 1;
			pivotTable.PivotFields("Не учитывать (=1)").ClearAllFilters();
			pivotTable.PivotFields("Не учитывать (=1)").CurrentPage = "(blank)";

			wsPivote.Activate();
			wsPivote.Columns["B:N"].Select();
			xlApp.Selection.ColumnWidth = 10;

			wsPivote.Range["A1"].Select();

			pivotTable.HasAutoFormat = false;
			pivotTable.ShowTableStyleColumnStripes = true;
			pivotTable.TableStyle2 = "PivotStyleMedium2";

			wsPivote.Columns["B:N"].Select();
			xlApp.Selection.FormatConditions.AddColorScale(3);
			xlApp.Selection.FormatConditions(xlApp.Selection.FormatConditions.Count).SetFirstPriority();

			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(1).Type = Excel.XlConditionValueTypes.xlConditionValueNumber;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(1).Value = 0;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor.TintAndShade = 0;

			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(2).Type = Excel.XlConditionValueTypes.xlConditionValueNumber;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 70;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = 13562593;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor.TintAndShade = 0;

			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(3).Type = Excel.XlConditionValueTypes.xlConditionValueNumber;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 150;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = 6870690;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor.TintAndShade = 0;

			wsPivote.Range["C1"].Select();

			wb.ShowPivotTableFieldList = false;
		}
	}
}
