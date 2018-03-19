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

				int necessaryServicesInMes = (from x in treat.DictMES where x.Value == 0 select x).Count();
				int hasAtLeastOneReferralByMes = treat.ListReferralsFromMes.Count > 0 ? 1 : 0;
				int necessaryServiceReferralByMesInstrumental = 0;
				int necessaryServiceReferralByMesLaboratory = 0;
				int necessaryServiceReferralCompletedByMes = 0;

				foreach (string item in treat.ListReferralsFromMes) {
					if (!treat.DictMES.ContainsKey(item))
						continue;

					if (treat.DictMES[item] == 0) {
						if (!treat.DictAllReferrals.ContainsKey(item))
							continue;

						if (treat.DictAllReferrals[item].RefType == 2)
							necessaryServiceReferralByMesLaboratory++;
						else
							necessaryServiceReferralByMesInstrumental++;
						
						if (treat.DictAllReferrals[item].IsCompleted == 1)
							necessaryServiceReferralCompletedByMes++;
					}
				}

				int hasAtLeastOneReferralSelfMade = (treat.DictAllReferrals.Count - treat.ListReferralsFromMes.Count) > 0 ? 1 : 0;
				int necessaryServiceReferralSelfMadeInstrumental = 0;
				int necessaryServiceReferralSelfMadeLaboratory = 0;
				int necessaryServiceReferralCompletedSelfMade = 0;

				foreach (string item in treat.ListReferralsFromDoc) {
					if (!treat.DictMES.ContainsKey(item))
						continue;

					if (treat.DictMES[item] == 0) {
						if (!treat.DictAllReferrals.ContainsKey(item))
							continue;

						if (treat.DictAllReferrals[item].RefType == 2)
							necessaryServiceReferralSelfMadeLaboratory++;
						else
							necessaryServiceReferralSelfMadeInstrumental++;

						if (treat.DictAllReferrals[item].IsCompleted == 1)
							necessaryServiceReferralCompletedSelfMade++;
					}
				}

				int servicesAllReferralsInstrumental = (from x in treat.DictAllReferrals where x.Value.RefType != 2 select x).Count();
				int servicesAllReferralsLaboratory = treat.DictAllReferrals.Count - servicesAllReferralsInstrumental;
				int completedServicesInReferrals = (from x in treat.DictAllReferrals where x.Value.IsCompleted == 1 select x).Count();
				int serviceInReferralOutsideMes = 0;
				foreach (KeyValuePair<string, ItemMESUsageTreatment.ReferralDetails> pair in treat.DictAllReferrals)
					if (!treat.DictMES.ContainsKey(pair.Key))
						serviceInReferralOutsideMes++;

				double necessaryServiceInMesUsedPercent =
					(double)(
					necessaryServiceReferralByMesInstrumental + 
					necessaryServiceReferralByMesLaboratory +
					necessaryServiceReferralSelfMadeInstrumental +
					necessaryServiceReferralSelfMadeLaboratory) / 
					(double)necessaryServicesInMes;
				
				List<object> values = new List<object>() {
					treatment.Key, //Код лечения
					1, //Прием
					treat.TREATDATE, //Дата лечения
					treat.FILIAL, //Филиал
					treat.DEPNAME, //Подразделение
					treat.DOCNAME, //ФИО врача
					treat.HISTNUM, //Номер ИБ
					treat.CLIENTNAME, //ФИО пациента
					treat.AGE, //Возраст
					treat.MKBCODE, //Код МКБ
					necessaryServicesInMes, //Кол-во обязательных услуг согласно МЭС
					treat.DictMES.Count, //Всего услуг в МЭС
					hasAtLeastOneReferralByMes, //Есть направление, созданное с использованием МЭС
					necessaryServiceReferralByMesInstrumental, //Кол-во обязательных услуг в направлении с использованием МЭС (инструментальных)
					necessaryServiceReferralByMesLaboratory, //Кол-во обязательных услуг в направлении с использованием МЭС (лабораторных)
					necessaryServiceReferralCompletedByMes, //Кол-во исполненных обязательных услуг в направлении МЭС
					hasAtLeastOneReferralSelfMade, //Есть направление, созданное самостоятельно
					necessaryServiceReferralSelfMadeInstrumental, //Кол-во обязательных услуг в направлении выставленных самостоятельно (инструментальных)
					necessaryServiceReferralSelfMadeLaboratory, //Кол-во обязательных услуг в направлении выставленных самостоятельно (лабораторных)
					necessaryServiceReferralCompletedSelfMade, //Кол-во исполненных обязательных услуг в самостоятельно созданных направлениях
					servicesAllReferralsInstrumental, //Всего услуг во всех направлениях (иснтрументальных)
					servicesAllReferralsLaboratory, //Всего услуг во всех направлениях (лабораторных)
					completedServicesInReferrals, //Кол-во выполненных услуг во всех направлениях
					serviceInReferralOutsideMes, //Кол-во услуг в направлениях, не входящих в МЭС
					necessaryServiceInMesUsedPercent, //% Соответствия обязательных услуг МЭС (обязательные во всех направлениях) / всего обязательных в мэс
					necessaryServiceInMesUsedPercent == 1 ? 1 : 0, //Услуги из всех направлений соответсвуют обязательным услугам МЭС на 100%
					treat.SERVICE_TYPE, //Тип приема
					treat.PAYMENT_TYPE //Тип оплаты приема
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

				for (int i = 2; i <= rowsUsed; i++)
					ws.Range["F" + i].FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)";

				ws.Columns["F:F"].Select();
				xlApp.Selection.NumberFormat = "0,0%";

				ws.Range["A" + (rowsUsed + 2)].Value = "Итого:";

				foreach (string item in new string[] { "B", "C", "D", "E" }) 
					ws.Range[item + (rowsUsed + 2)].Formula = "=SUM(" + item + "2:" + item + rowsUsed + ")";

				ws.Range["F" + (rowsUsed + 2)].FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)";


				//string rangeData = "A1:A" + rowsUsed + ",F1:F" + rowsUsed;
				//Console.WriteLine("rangeData: " + rangeData);
				xlApp.ActiveSheet.Shapes.AddChart2(201, Excel.XlChartType.xlColumnClustered).Select();
				xlApp.ActiveChart.SetSourceData(ws.get_Range("A1:A2;F1:F2"));
				xlApp.ActiveSheet.Shapes["Диаграмма 1"].Top = 0;
				xlApp.ActiveSheet.Shapes["Диаграмма 1"].Left = 480;

				//rowsUsed += 2;
				//ws.Range["A" + rowsUsed].Value = "Итого:";
				//ws.Range["B" + rowsUsed].Formula = "=AVERAGE(B2:B" + (rowsUsed - 2) + ")";
				//ws.Range["A" + rowsUsed + ":B" + rowsUsed].Select();
				//xlApp.Selection.Interior.Pattern = Excel.Constants.xlSolid;
				//xlApp.Selection.Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
				//xlApp.Selection.Interior.Color = 65535;
				//xlApp.Selection.Interior.TintAndShade = 0;
				//xlApp.Selection.Font.Bold = Excel.Constants.xlSolid;
				//rowsUsed++;
				ws.Range["A" + rowsUsed].Select();
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
