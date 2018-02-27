using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports {
	class NpoiExcelGeneral {
		private static bool CreateNewIWorkbook(string resultFilePrefix, string templateFileName, out IWorkbook workbook, out ISheet sheet, out string resultFile) {
			workbook = null;
			sheet = null;
			resultFile = string.Empty;

			try {
				string templateFile = Program.AssemblyDirectory + templateFileName;
				foreach (char item in Path.GetInvalidFileNameChars())
					resultFilePrefix = resultFilePrefix.Replace(item, '-');

				if (!File.Exists(templateFile)) {
					Logging.ToFile("Не удалось найти файл шаблона: " + templateFile);
					return false;
				}

				string resultPath = Path.Combine(Program.AssemblyDirectory, "Results");
				if (!Directory.Exists(resultPath))
					Directory.CreateDirectory(resultPath);

				resultFile = Path.Combine(resultPath, resultFilePrefix + ".xlsx");

				using (FileStream stream = new FileStream(templateFile, FileMode.Open, FileAccess.Read))
					workbook = new XSSFWorkbook(stream);

				sheet = workbook.GetSheet("Данные");

				return true;
			} catch (Exception e) {
				Logging.ToFile(e.Message + Environment.NewLine + e.StackTrace);
				return false;
			}
		}

		private static bool SaveAndCloseIWorkbook(IWorkbook workbook, string resultFile) {
			try {
				using (FileStream stream = new FileStream(resultFile, FileMode.Create, FileAccess.Write))
					workbook.Write(stream);

				workbook.Close();

				return true;
			} catch (Exception e) {
				Logging.ToFile(e.Message + Environment.NewLine + e.StackTrace);
				return false;
			}
		}


		public static string WriteDataTableToExcel(DataTable dataTable, string resultFilePrefix, string templateFileName) {
			if (!CreateNewIWorkbook(resultFilePrefix, templateFileName, out IWorkbook workbook, out ISheet sheet, out string resultFile))
				return string.Empty;

			int rowNumber = 1;
			int columnNumber = 0;

			foreach (DataRow dataRow in dataTable.Rows) {
				IRow row = sheet.CreateRow(rowNumber);

				foreach (DataColumn column in dataTable.Columns) {
					ICell cell = row.CreateCell(columnNumber);
					string value = dataRow[column].ToString();

					if (double.TryParse(value, out double result)) {
						cell.SetCellValue(result);
					} else if (DateTime.TryParseExact(value, "dd.MM.yyyy h:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date)) {
						cell.SetCellValue(date);
					} else {
						cell.SetCellValue(value);
					}

					columnNumber++;
				}

				columnNumber = 0;
				rowNumber++;
			}

			if (!SaveAndCloseIWorkbook(workbook, resultFile))
				return string.Empty;

			return resultFile;
		}



		public static string WriteMesUsageTreatmentsToExcel(Dictionary<string, ItemMESUsageTreatment> treatments, string resultFilePrefix, string templateFileName) {
			if (!CreateNewIWorkbook(resultFilePrefix, templateFileName, out IWorkbook workbook, out ISheet sheet, out string resultFile))
				return string.Empty;

			int rowNumber = 1;
			int columnNumber = 0;

			foreach (KeyValuePair<string, ItemMESUsageTreatment> treatment in treatments) {
				IRow row = sheet.CreateRow(rowNumber);

				ItemMESUsageTreatment treat = treatment.Value;
				double percentCompleted =
					((double)treat.ListReferralsFromMes.Count +
					(double)treat.ListReferralsFromDoc.Count) /
					(double)treat.ListMES.Count;

				int mesReferralsExecuted = 0;
				int docReferralsExecuted = 0;
				int allReferralsExecuted = 0;
				int oversizedReferral = 0;

				foreach (string item in treat.ListReferralsFromMes) {
					if (!treat.ListAllReferrals.ContainsKey(item))
						continue;

					mesReferralsExecuted += treat.ListAllReferrals[item];
				}

				foreach (string item in treat.ListReferralsFromDoc) {
					if (!treat.ListAllReferrals.ContainsKey(item))
						continue;

					docReferralsExecuted += treat.ListAllReferrals[item];
				}

				foreach (KeyValuePair<string, int> pair in treat.ListAllReferrals) {
					allReferralsExecuted += pair.Value;

					if (!treat.ListMES.Contains(pair.Key))
						oversizedReferral++;
				}
				
				List<object> values = new List<object>() {
					treatment.Key,
					1,
					treat.TREATDATE,
					treat.FILIAL,
					treat.DEPNAME,
					treat.DOCNAME,
					treat.HISTNUM,
					treat.CLIENTNAME,
					treat.AGE,
					treat.MKBCODE,
					treat.ListMES.Count,
					treat.ListReferralsFromMes.Count > 0 ? 1 : 0,
					treat.ListReferralsFromMes.Count,
					mesReferralsExecuted,
					treat.ListReferralsFromDoc.Count > 0 ? 1 : 0,
					treat.ListReferralsFromDoc.Count,
					docReferralsExecuted,
					treat.ListAllReferrals.Count > 0 ? 1 : 0,
					treat.ListAllReferrals.Count,
					allReferralsExecuted,
					oversizedReferral,
					percentCompleted,
					percentCompleted == 1 ? 1 : 0
				};

				foreach (object value in values) {
					ICell cell = row.CreateCell(columnNumber);

					if (double.TryParse(value.ToString(), out double result))
						cell.SetCellValue(result);
					else
						cell.SetCellValue(value.ToString());

					columnNumber++;
				}

				columnNumber = 0;
				rowNumber++;
			}
			
			if (!SaveAndCloseIWorkbook(workbook, resultFile))
				return string.Empty;

			return resultFile;
		}

		public static bool PerformOnlineAccountsUsage(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, 
				out Excel.Worksheet ws))
				return false;

			try {
				int rowsUsed = ws.UsedRange.Rows.Count;
				ws.Columns["B:B"].Select();
				xlApp.Selection.NumberFormat = "0,0%";
				xlApp.ActiveSheet.Shapes.AddChart2(201, Excel.XlChartType.xlColumnClustered).Select();
				xlApp.ActiveChart.SetSourceData(ws.Range["A1:B" + rowsUsed]);
				//xlApp.ActiveSheet.Shapes["Диаграмма 1"].IncrementLeft(-237);
				xlApp.ActiveSheet.Shapes["Диаграмма 1"].IncrementTop(-174);
				ws.Range["A" + (rowsUsed + 1)].Select();
			} catch (Exception e) {
				Logging.ToFile(e.Message + Environment.NewLine + e.StackTrace);
			}
			
			SaveAndCloseWorkbook(xlApp, wb);

			return true;
		}
		
		public static bool PerformFreeCells(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, 
				out Excel.Worksheet ws))
				return false;

			try {
				ws.Columns["B:B"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ";
				ws.Columns["B:B"].EntireColumn.AutoFit();
			} catch (Exception e) {
				Logging.ToFile(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				AddPivotTableFreeCells(wb, ws, xlApp);
			} catch (Exception e) {
				Logging.ToFile(e.Message + Environment.NewLine + e.StackTrace);
			}

			wb.Sheets["Сводная таблица"].Activate();
			SaveAndCloseWorkbook(xlApp, wb);

			return true;
		}

		private static void AddPivotTableFreeCells(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			ws.Cells[1, 1].Select();

			string pivotTableName = @"PivotTable";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная таблица"];
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


			pivotTable.AddDataField(pivotTable.PivotFields("Всего"), "(Всего)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.AddDataField(pivotTable.PivotFields("Занято"), "(Занято)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.AddDataField(pivotTable.PivotFields("Загрузка"), "(Загрузка)", Excel.XlConsolidationFunction.xlAverage);

			pivotTable.PivotFields("Дата").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
			pivotTable.PivotFields("Дата").Position = 1;

			pivotTable.PivotFields("Дата").AutoGroup();

			pivotTable.PivotFields("Филиал").PivotItems("Call-центр").Visible = false;
			pivotTable.PivotFields("Филиал").PivotItems("КУТУЗ").Visible = false;
			pivotTable.PivotFields("Филиал").PivotItems("СКОРАЯ").Visible = false;

			pivotTable.RowGrand = false;
			pivotTable.ColumnGrand = false;
			pivotTable.DisplayFieldCaptions = false;

			pivotTable.PivotFields("(Занято)").NumberFormat = "0,00";
			pivotTable.PivotFields("(Загрузка)").NumberFormat = "0,0%";
			
			pivotTable.PivotSelect("'(Загрузка)'", Excel.XlPTSelectionMode.xlDataAndLabel, true);

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
			
			
			pivotTable.PivotFields("Филиал").AutoSort(Excel.XlSortOrder.xlAscending, "(Загрузка)");

			pivotTable.PivotFields("Отделение").ShowDetail = false;
			pivotTable.PivotFields("Пересечение").ShowDetail = false;
			pivotTable.PivotFields("Филиал").ShowDetail = false;

			try {
				pivotTable.PivotFields("Месяцы").Orientation = Excel.XlPivotFieldOrientation.xlHidden;
			} catch (Exception) {
			}

			wsPivote.Range["A1"].Select();
			wb.ShowPivotTableFieldList = false;
		}
		
		private static bool OpenWorkbook(string workbook, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws) {
			xlApp = null;
			wb = null;
			ws = null;

			xlApp = new Excel.Application();

			if (xlApp == null) {
				Logging.ToFile("Не удалось открыть приложение Excel");
				return false;
			}

			xlApp.Visible = false;

			wb = xlApp.Workbooks.Open(workbook);

			if (wb == null) {
				Logging.ToFile("Не удалось открыть книгу " + workbook);
				return false;
			}

			ws = wb.Sheets["Данные"];

			if (ws == null) {
				Logging.ToFile("Не удалось открыть лист Данные");
				return false;
			}

			return true;
		}

		private static void SaveAndCloseWorkbook(Excel.Application xlApp, Excel.Workbook wb) {
			wb.Save();
			wb.Close();

			xlApp.Quit();
		}
		
		public static bool PerformUnclosedProtocols(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws))
				return false;

			try {
				int usedRows = ws.UsedRange.Rows.Count;
				ws.Range["A1"].Select();
				xlApp.Selection.AutoFilter();
				ws.Columns["B:B"].Select();
				xlApp.Selection.Insert(Excel.XlDirection.xlToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
				ws.Range["B1"].Value = "Уникальное лечение";
				ws.Range["B2"].Select();
				xlApp.ActiveCell.FormulaR1C1 = "=IF(RC[-1]=R[1]C[-1],0,1)";
				xlApp.Selection.AutoFill(ws.Range["B2:B" + usedRows]);
				ws.Columns["F:F"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ";
			} catch (Exception e) {
				Logging.ToFile(e.Message + Environment.NewLine + e.StackTrace);
			}
			
			try {
				UnclosedProtocolsAddPivotTableDepartments(wb, ws, xlApp);
			} catch (Exception e) {
				Logging.ToFile(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				UnclosedProtocolsAddPivotTableDoctors(wb, ws, xlApp);
			} catch (Exception e) {
				Logging.ToFile(e.Message + Environment.NewLine + e.StackTrace);
			}

			wb.Sheets["Сводная по отделениям"].Activate();
			SaveAndCloseWorkbook(xlApp, wb);

			return true;
		}

		private static void UnclosedProtocolsAddPivotTableDoctors(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			ws.Cells[1, 1].Select();

			string pivotTableName = @"WorkTimePivotTable";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная по врачам"];

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			pivotTable.PivotFields("ФИО врача").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ФИО врача").Position = 1;

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 2;

			pivotTable.PivotFields("Подразделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Подразделение").Position = 3;

			pivotTable.PivotFields("ФИО врача").Subtotals = 
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("ФИО врача").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Филиал").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Филиал").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Подразделение").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Подразделение").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.AddDataField(pivotTable.PivotFields("Уникальное лечение"), "Кол-во лечений", Excel.XlConsolidationFunction.xlSum);

			pivotTable.CalculatedFields().Add("Кол-во протоколов", "='Протокол подписан' +'Протокол не подписан'", true);
			pivotTable.PivotFields("Кол-во протоколов").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю Кол-во протоколов").Caption = "Общее кол-во протоколов";

			pivotTable.AddDataField(pivotTable.PivotFields("Протокол не подписан"), "Кол-во неподписанных", Excel.XlConsolidationFunction.xlSum);

			pivotTable.CalculatedFields().Add("Процент неподписанных", "='Протокол не подписан' /'Кол-во протоколов'", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Процент неподписанных").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю Процент неподписанных").Caption = "Доля неподписанных протоколов";
			pivotTable.PivotFields("Доля неподписанных протоколов").NumberFormat = "0,00%";
			
			pivotTable.PivotFields("ФИО врача").AutoSort(Excel.XlSortOrder.xlDescending,
				"Доля неподписанных протоколов");

			pivotTable.HasAutoFormat = false;
			wb.ShowPivotTableFieldList = false;
		}

		private static void UnclosedProtocolsAddPivotTableDepartments(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			ws.Cells[1, 1].Select();

			string pivotTableName = @"WorkTimePivotTable";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная по отделениям"];

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 1;

			pivotTable.PivotFields("Подразделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Подразделение").Position = 2;

			pivotTable.PivotFields("ФИО врача").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ФИО врача").Position = 3;

			pivotTable.AddDataField(pivotTable.PivotFields("Уникальное лечение"), "Кол-во лечений", Excel.XlConsolidationFunction.xlSum);

			pivotTable.CalculatedFields().Add("Всего протоколов", "='Протокол подписан' +'Протокол не подписан'", true);
			pivotTable.PivotFields("Всего протоколов").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю Всего протоколов").Caption = "Общее кол-во протоколов";

			pivotTable.AddDataField(pivotTable.PivotFields("Протокол не подписан"), "Кол-во неподписанных протоколов", Excel.XlConsolidationFunction.xlSum);

			pivotTable.CalculatedFields().Add("Процент неподписанных", "='Протокол не подписан' /'Всего протоколов'", true);
			pivotTable.PivotFields("Процент неподписанных").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю Процент неподписанных").Caption = "Доля неподписанных протоколов";
			pivotTable.PivotFields("Доля неподписанных протоколов").NumberFormat = "0,00%";

			pivotTable.PivotFields("Филиал").AutoSort(Excel.XlSortOrder.xlDescending, "Доля неподписанных протоколов");
			pivotTable.PivotFields("Подразделение").AutoSort(Excel.XlSortOrder.xlDescending, "Доля неподписанных протоколов");
			pivotTable.PivotFields("ФИО врача").AutoSort(Excel.XlSortOrder.xlDescending, "Доля неподписанных протоколов");

			pivotTable.PivotFields("Подразделение").ShowDetail = false;
			pivotTable.PivotFields("Филиал").ShowDetail = false;

			pivotTable.HasAutoFormat = false;

			wsPivote.Columns[1].ColumnWidth = 60;
			wb.ShowPivotTableFieldList = false;
		}
	}
}
