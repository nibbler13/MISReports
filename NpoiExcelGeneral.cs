using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports {
	class NpoiExcelGeneral {
        private static bool CreateNewIWorkbook(string resultFilePrefix, string templateFileName,
			out IWorkbook workbook, out ISheet sheet, out string resultFile, string sheetName) {
			workbook = null;
			sheet = null;
			resultFile = string.Empty;

			try {
				if (!GetTemplateFilePath(ref templateFileName))
					return false;

				string resultPath = GetResultFilePath(resultFilePrefix, templateFileName);

				using (FileStream stream = new FileStream(templateFileName, FileMode.Open, FileAccess.Read))
					workbook = new XSSFWorkbook(stream);

				if (string.IsNullOrEmpty(sheetName))
					sheetName = "Данные";

				sheet = workbook.GetSheet(sheetName);
				resultFile = resultPath;

				return true;
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				return false;
			}
		}

		private static bool GetTemplateFilePath(ref string templateFileName) {
			templateFileName = Path.Combine(Program.AssemblyDirectory, templateFileName);

			if (!File.Exists(templateFileName)) {
				Logging.ToLog("Не удалось найти файл шаблона: " + templateFileName);
				return false;
			}

			return true;
		}

		private static string GetResultFilePath(string resultFilePrefix, string templateFileName, bool isPlainText = false) {
			string resultPath = Path.Combine(Program.AssemblyDirectory, "Results");
			if (!Directory.Exists(resultPath))
				Directory.CreateDirectory(resultPath);

			foreach (char item in Path.GetInvalidFileNameChars())
				resultFilePrefix = resultFilePrefix.Replace(item, '-');

			string fileEnding = ".xlsx";
			if (isPlainText)
				fileEnding = ".txt";

			string resultFile = Path.Combine(resultPath, resultFilePrefix + " " + DateTime.Now.ToString("yyyyMMdd_HHmmss") + fileEnding);
			if (isPlainText)
				File.Copy(templateFileName, resultFile, true);

			return resultFile;
		}

		private static bool SaveAndCloseIWorkbook(IWorkbook workbook, string resultFile) {
			try {
				using (FileStream stream = new FileStream(resultFile, FileMode.Create, FileAccess.Write))
					workbook.Write(stream);

				workbook.Close();

				return true;
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				return false;
			}
		}



		private static bool OpenWorkbook(string workbook, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws, string sheetName = "") {
			xlApp = null;
			wb = null;
			ws = null;

			xlApp = new Excel.Application();

			if (xlApp == null) {
				Logging.ToLog("Не удалось открыть приложение Excel");
				return false;
			}

			xlApp.Visible = false;

			wb = xlApp.Workbooks.Open(workbook);

			if (wb == null) {
				Logging.ToLog("Не удалось открыть книгу " + workbook);
				return false;
			}

			if (string.IsNullOrEmpty(sheetName))
				sheetName = "Данные";

			ws = wb.Sheets[sheetName];

			if (ws == null) {
				Logging.ToLog("Не удалось открыть лист Данные");
				return false;
			}

			return true;
		}

        private static void SaveAndCloseWorkbook(Excel.Application xlApp, Excel.Workbook wb, Excel.Worksheet ws) {
			if (ws != null) {
                Marshal.ReleaseComObject(ws);
                ws = null;
            }

			if (wb != null) {
				wb.Save();
				wb.Close(0);
				Marshal.ReleaseComObject(wb);
                wb = null;
			}

			if (xlApp != null) {
				xlApp.Quit();
				Marshal.ReleaseComObject(xlApp);
                xlApp = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
		}



		public static string WriteDataTableToExcel(DataTable dataTable, string resultFilePrefix, string templateFileName,
			bool telemedicineOnlyIngosstrakh = false, string sheetName = "", bool createNew = true) {
			IWorkbook workbook = null;
			ISheet sheet = null;
			string resultFile = string.Empty;

			if (createNew) {
				if (!CreateNewIWorkbook(resultFilePrefix, templateFileName,
					out workbook, out sheet, out resultFile, sheetName))
					return string.Empty;
			} else {
				try {
					using (FileStream stream = new FileStream(templateFileName, FileMode.Open, FileAccess.Read))
						workbook = new XSSFWorkbook(stream);

					sheet = workbook.GetSheet(sheetName);
					resultFile = templateFileName;
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					return string.Empty;
				}
			}

			int rowNumber = 1;
			int columnNumber = 0;

			foreach (DataRow dataRow in dataTable.Rows) {
				IRow row = null;
				try { row = sheet.GetRow(rowNumber); } catch (Exception) { }

				if (row == null)
					row = sheet.CreateRow(rowNumber);

				if (telemedicineOnlyIngosstrakh) {
					try {
						string paymentType = dataRow["JNAME"].ToString();
						if (!paymentType.ToLower().Contains("ингосстрах"))
							continue;
					} catch (Exception) { }
				}

				foreach (DataColumn column in dataTable.Columns) {
					ICell cell = null;
					try { cell = row.GetCell(columnNumber); } catch (Exception) { }

					if (cell == null)
						cell = row.CreateCell(columnNumber);

					string value = dataRow[column].ToString();

					if (double.TryParse(value, out double result)) {
						cell.SetCellValue(result);
					} else if (DateTime.TryParse(value, out DateTime date)) {
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
		
		public static string WriteDataTableToTextFile(DataTable dataTable, string resultFilePrefix, string templateFileName) {
			string resultFile = string.Empty;

			try {
				if (!GetTemplateFilePath(ref templateFileName))
					return resultFile;

				resultFile = GetResultFilePath(resultFilePrefix, templateFileName, true);

				using (System.IO.StreamWriter sw = System.IO.File.AppendText(resultFile)) {
					foreach (DataRow dataRow in dataTable.Rows) {
						object[] values = dataRow.ItemArray;
						List<string> valuesToWrite = new List<string>();
						foreach (object value in values)
							valuesToWrite.Add(value.ToString().Replace(" 0:00:00", ""));

						if (valuesToWrite.Count > 0) {
							string logLine = string.Join("	", valuesToWrite.ToArray());
							sw.WriteLine();
							sw.Write(logLine);
						}
					}
				}
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			return resultFile;
		}
			   		 
		public static string WriteMesUsageTreatmentsToExcel(Dictionary<string, ItemMESUsageTreatment> treatments, string resultFilePrefix, string templateFileName) {
			if (!CreateNewIWorkbook(resultFilePrefix, templateFileName, out IWorkbook workbook, out ISheet sheet, out string resultFile, string.Empty))
				return string.Empty;

			int rowNumber = 1;
			int columnNumber = 0;

			foreach (KeyValuePair<string, ItemMESUsageTreatment> treatment in treatments) {
				IRow row = sheet.CreateRow(rowNumber);
				ItemMESUsageTreatment treat = treatment.Value;

				int necessaryServicesInMes = (from x in treat.DictMES where x.Value == 0 select x).Count();

				if (necessaryServicesInMes == 0)
					continue;

				int hasAtLeastOneReferralByMes = treat.ListReferralsFromMes.Count > 0 ? 1 : 0;
				int necessaryServiceReferralByMesInstrumental = 0;
				int necessaryServiceReferralByMesLaboratory = 0;
				int necessaryServiceReferralCompletedByMesInstrumental = 0;
				int necessaryServiceReferralCompletedByMesLaboratory = 0;

				foreach (string item in treat.ListReferralsFromMes) {
					if (!treat.DictMES.ContainsKey(item))
						continue;

					if (treat.DictMES[item] == 0) {
						if (!treat.DictAllReferrals.ContainsKey(item))
							continue;

						int isCompleted = treat.DictAllReferrals[item].IsCompleted == 1 ? 1 : 0;

						int refType = treat.DictAllReferrals[item].RefType;
						if (refType == 2 || refType == 992140066) {
							necessaryServiceReferralByMesLaboratory++;
							necessaryServiceReferralCompletedByMesLaboratory += isCompleted;
						} else {
							necessaryServiceReferralByMesInstrumental++;
							necessaryServiceReferralCompletedByMesInstrumental += isCompleted;
						}
					}
				}

				int hasAtLeastOneReferralSelfMade = (treat.DictAllReferrals.Count - treat.ListReferralsFromMes.Count) > 0 ? 1 : 0;
				int necessaryServiceReferralSelfMadeInstrumental = 0;
				int necessaryServiceReferralSelfMadeLaboratory = 0;
				int necessaryServiceReferralCompletedSelfMadeInstrumental = 0;
				int necessaryServiceReferralCompletedSelfMadeLaboratory = 0;

				foreach (string item in treat.ListReferralsFromDoc) {
					if (!treat.DictMES.ContainsKey(item))
						continue;

					if (treat.DictMES[item] == 0) {
						if (!treat.DictAllReferrals.ContainsKey(item))
							continue;

						int isCompleted = treat.DictAllReferrals[item].IsCompleted == 1 ? 1 : 0;

						int refType = treat.DictAllReferrals[item].RefType;
						if (refType == 2 || refType == 992140066) {
							necessaryServiceReferralSelfMadeLaboratory++;
							necessaryServiceReferralCompletedSelfMadeLaboratory += isCompleted;
						} else {
							necessaryServiceReferralSelfMadeInstrumental++;
							necessaryServiceReferralCompletedSelfMadeInstrumental += isCompleted;
						}
					}
				}

				int servicesAllReferralsInstrumental = (from x in treat.DictAllReferrals where x.Value.RefType != 2 select x).Count();
				int servicesAllReferralsLaboratory = treat.DictAllReferrals.Count - servicesAllReferralsInstrumental;
				int completedServicesInReferrals = (from x in treat.DictAllReferrals where x.Value.IsCompleted == 1 select x).Count();
				int serviceInReferralOutsideMes = 0;
				foreach (KeyValuePair<string, ItemMESUsageTreatment.ReferralDetails> pair in treat.DictAllReferrals)
					if (!treat.DictMES.ContainsKey(pair.Key))
						serviceInReferralOutsideMes++;

				double necessaryServiceInMesUsedPercent;
				if (necessaryServicesInMes > 0)
					necessaryServiceInMesUsedPercent =
					(double)(
					necessaryServiceReferralByMesInstrumental +
					necessaryServiceReferralByMesLaboratory +
					necessaryServiceReferralSelfMadeInstrumental +
					necessaryServiceReferralSelfMadeLaboratory) /
					(double)necessaryServicesInMes;
				else
					necessaryServiceInMesUsedPercent = 1;

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
					//treat.DictMES.Count, //Всего услуг в МЭС
					hasAtLeastOneReferralByMes, //Есть направление, созданное с использованием МЭС
					necessaryServiceReferralByMesInstrumental + necessaryServiceReferralByMesLaboratory, //Кол-во услуг в направлении с использованием МЭС
					//necessaryServiceReferralByMesInstrumental, //Кол-во обязательных услуг в направлении с использованием МЭС (инструментальных)
					//necessaryServiceReferralByMesLaboratory, //Кол-во обязательных услуг в направлении с использованием МЭС (лабораторных)
					//necessaryServiceReferralCompletedByMesInstrumental, //Кол-во исполненных обязательных услуг в направлении МЭС (инструментальных)
					//necessaryServiceReferralCompletedByMesLaboratory, //Кол-во исполненных обязательных услуг в направлении МЭС (лабораторных)
					hasAtLeastOneReferralSelfMade, //Есть направление, созданное самостоятельно
					necessaryServiceReferralSelfMadeInstrumental + necessaryServiceReferralSelfMadeLaboratory, //Кол-во услуг в направлении выставленных самостоятельно
					//necessaryServiceReferralSelfMadeInstrumental, //Кол-во обязательных услуг в направлении выставленных самостоятельно (инструментальных)
					//necessaryServiceReferralSelfMadeLaboratory, //Кол-во обязательных услуг в направлении выставленных самостоятельно (лабораторных)
					//necessaryServiceReferralCompletedSelfMadeInstrumental, //Кол-во исполненных обязательных услуг в самостоятельно созданных направлениях (инструментальных)
					//necessaryServiceReferralCompletedSelfMadeLaboratory, //Кол-во исполненных обязательных услуг в самостоятельно созданных направлениях (лабораторных)
					//servicesAllReferralsInstrumental, //Всего услуг во всех направлениях (иснтрументальных)
					//servicesAllReferralsLaboratory, //Всего услуг во всех направлениях (лабораторных)
					//completedServicesInReferrals, //Кол-во выполненных услуг во всех направлениях
					//serviceInReferralOutsideMes, //Кол-во услуг в направлениях, не входящих в МЭС
					necessaryServiceInMesUsedPercent, //% Соответствия обязательных услуг МЭС (обязательные во всех направлениях) / всего обязательных в мэс
					necessaryServiceInMesUsedPercent == 1 ? 1 : 0, //Услуги из всех направлений соответсвуют обязательным услугам МЭС на 100%
					treat.SERVICE_TYPE, //Тип приема
					treat.PAYMENT_TYPE//, //Тип оплаты приема
					//treat.AGNAME, //Наименование организации
					//treat.AGNUM //Номер договора
				};

				foreach (object value in values) {
					ICell cell = row.CreateCell(columnNumber);

					if (double.TryParse(value.ToString(), out double result))
						cell.SetCellValue(result);
					else if (DateTime.TryParseExact(value.ToString(), "dd.MM.yyyy h:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
						cell.SetCellValue(date);
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



		public static bool PerformMesUsage(string resultFile) {
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

		private static void AddInteriorColor(Excel.Range range, Excel.XlThemeColor xlThemeColor) {
			range.Interior.Pattern = Excel.Constants.xlSolid;
			range.Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
			range.Interior.ThemeColor = xlThemeColor;
			range.Interior.TintAndShade = 0.799981688894314;
			range.Interior.PatternTintAndShade = 0;
		}



		public static bool PerformWorkload(string resultFile) {
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
					"ДЕЖУРНЫЙ ВРАЧ",
					"ПРОЦЕДУРНЫЙ",
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
							!filialCode.Equals("15"))
							ws.Range["AL" + row].Value2 = 1;

					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					}
				}

				ws.Columns["AM:AM"].Select();
				xlApp.Selection.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1:"=ДЛСТР(СЖПРОБЕЛЫ(AM1))=0");
				xlApp.Selection.FormatConditions(xlApp.Selection.FormatConditions.Count).SetFirstPriority();
				xlApp.Selection.FormatConditions(1).Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
				xlApp.Selection.FormatConditions(1).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
				xlApp.Selection.FormatConditions(1).Interior.TintAndShade = 0;

				xlApp.Selection.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, "=20");
				xlApp.Selection.FormatConditions(2).Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
				xlApp.Selection.FormatConditions(2).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
				xlApp.Selection.FormatConditions(2).Interior.TintAndShade = 0.799981688894314;

				xlApp.Selection.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, "=150");
				xlApp.Selection.FormatConditions(3).Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
				xlApp.Selection.FormatConditions(3).Interior.Color = 65535;
				xlApp.Selection.FormatConditions(3).Interior.TintAndShade = 0;

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



		public static bool PerformNonAppearance(string resultFile, DataTable dataTable) {
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
						Convert.ToInt32(row["WITHOUT_MARKS_WITHOUT_TREATMENTS"].ToString());

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



		public static bool PerformTelemedicine(string resultFile) {
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



		public static bool PerformVIP(string resultFile, string previousFile) {
			Logging.ToLog("Подготовка файла с отчетом по VIP-пациентам: " + resultFile);
			Logging.ToLog("Предыдущий файл: " + previousFile);
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws)) {
				Logging.ToLog("Не удалось открыть книгу: " + resultFile);
				return false;
			}

			try {
				ws.Columns["B:B"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ";
				ws.Columns["C:D"].Select();
				xlApp.Selection.NumberFormat = "ч:мм;@";
				ws.Columns["K:K"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ";
				ws.Cells[1, 1].Select();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}
			
			SaveAndCloseWorkbook(xlApp, wb, ws);

			if (string.IsNullOrEmpty(previousFile) || !File.Exists(previousFile)) {
				Logging.ToLog("Пропуск сравнения с предыдущей версией, файл не существует");
				return true;
			}

			Logging.ToLog("Считывание содержимого файлов");

			DataTable dataTableCurrent = ReadExcelFile(resultFile, "Данные");
			Logging.ToLog("Текущий файл, строк: " + dataTableCurrent.Rows.Count);

			DataTable dataTablePrevious = ReadExcelFile(previousFile, "Данные");
			Logging.ToLog("Предыдущий файл, строк: " + dataTablePrevious.Rows.Count);

			if (dataTablePrevious.Columns.Count == 14)
				dataTablePrevious.Columns.RemoveAt(13);

			if (!OpenWorkbook(resultFile, out xlApp, out wb, out ws)) {
				Logging.ToLog("Не удалось открыть книгу: " + resultFile);
				return false;
			}

			for (int i = 1; i < dataTableCurrent.Rows.Count; i++) {
				DataRow dataRowLeft = dataTableCurrent.Rows[i];
				bool existedBefore = false;

				for (int k = 1; k < dataTablePrevious.Rows.Count; k++) {
					DataRow dataRowRight = dataTablePrevious.Rows[k];
					if (DataRowComparer.Default.Equals(dataRowLeft, dataRowRight)) {
						existedBefore = true;
						break;
					}
				}

				if (!existedBefore) {
					int rowNumber = i + 1;
					ws.Range["A" + rowNumber + ":N" + rowNumber].Interior.ColorIndex = 35;
					ws.Range["N" + rowNumber + ":N" + rowNumber].Value2 = "Новая запись";
				}
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static DataTable ReadExcelFile(string fileName, string sheetName) {
			Logging.ToLog("Считывание файла: " + fileName + ", лист: " + sheetName);
			DataTable dataTable = new DataTable();

			if (!File.Exists(fileName))
				return dataTable;

			try {
				using (OleDbConnection conn = new OleDbConnection()) {
					conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Mode=Read;" +
						"Extended Properties='Excel 12.0 Xml;HDR=NO;'";

					using (OleDbCommand comm = new OleDbCommand()) {
						if (string.IsNullOrEmpty(sheetName)) {
							conn.Open();
							DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
								new object[] { null, null, null, "TABLE" });
							sheetName = dtSchema.Rows[0].Field<string>("TABLE_NAME");
							conn.Close();
						} else
							sheetName += "$";

						comm.CommandText = "Select * from [" + sheetName + "]";
						comm.Connection = conn;

						using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter()) {
							oleDbDataAdapter.SelectCommand = comm;
							oleDbDataAdapter.Fill(dataTable);
						}
					}
				}
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			return dataTable;
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
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}
			
			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}



		public static bool PerformFreeCells(string resultFile, DateTime dateBeginOriginal, DateTime dateEnd) {
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
				ws.Columns["G:G"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ";
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}
			
			try {
				UnclosedProtocolsAddPivotTableDepartments(wb, ws, xlApp);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				UnclosedProtocolsAddPivotTableDoctors(wb, ws, xlApp);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			wb.Sheets["Сводная по отделениям"].Activate();
			SaveAndCloseWorkbook(xlApp, wb, ws);

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
			pivotTable.PivotFields("ФИО врача").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("ФИО врача").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("DCODE").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("DCODE").Position = 2;
			pivotTable.PivotFields("DCODE").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("DCODE").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 3;
			pivotTable.PivotFields("Филиал").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Филиал").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Подразделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Подразделение").Position = 4;
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

			pivotTable.PivotFields("Статус сотрудника").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
			pivotTable.PivotFields("Статус сотрудника").Position = 1;

			wsPivote.Columns[2].ColumnWidth = 12;

			/*
			With ActiveSheet.PivotTables("WorkTimePivotTable").PivotFields("DCODE")
				.Orientation = xlRowField
				.Position = 2
			End With
			ActiveSheet.PivotTables("WorkTimePivotTable").PivotFields("DCODE").Subtotals = _
				Array(False, False, False, False, False, False, False, False, False, False, False, False)
			ActiveSheet.PivotTables("WorkTimePivotTable").PivotFields("DCODE").LayoutForm _
				= xlTabular
			 */

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
			pivotTable.PivotFields("ФИО врача").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("ФИО врача").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("DCODE").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("DCODE").Position = 4;
			pivotTable.PivotFields("DCODE").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("DCODE").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			/*
			 With ActiveSheet.PivotTables("WorkTimePivotTable").PivotFields("DCODE")
				.Orientation = xlRowField
				.Position = 4
			End With
			ActiveSheet.PivotTables("WorkTimePivotTable").PivotFields("DCODE").Subtotals = _
				Array(False, False, False, False, False, False, False, False, False, False, False, False)
			ActiveSheet.PivotTables("WorkTimePivotTable").PivotFields("DCODE").LayoutForm _
				= xlTabular
			ActiveSheet.PivotTables("WorkTimePivotTable").PivotFields("ФИО врача"). _
				Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
				False, False)
			ActiveSheet.PivotTables("WorkTimePivotTable").PivotFields("ФИО врача"). _
				LayoutForm = xlTabular
			 */

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

			pivotTable.PivotFields("Статус сотрудника").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
			pivotTable.PivotFields("Статус сотрудника").Position = 1;

			//With ActiveSheet.PivotTables("WorkTimePivotTable").PivotFields( _
			//	"Статус сотрудника")
			//	.Orientation = xlPageField
			//	.Position = 1
			//End With

			pivotTable.PivotFields("Подразделение").ShowDetail = false;
			pivotTable.PivotFields("Филиал").ShowDetail = false;

			pivotTable.HasAutoFormat = false;

			wsPivote.Columns[1].ColumnWidth = 60;
			wsPivote.Columns[2].ColumnWidth = 12;
			wb.ShowPivotTableFieldList = false;
		}



		public static bool PerformRegistryMarks(
            string resultFile, DataTable dataTable, DateTime dateTimeBegin) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws))
				return false;

			try {
				ws.Columns["C:C"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ Ч:мм;@";
				ws.Range["A1"].Select();
                xlApp.Selection.AutoFilter();
                ws.UsedRange.AutoFilter(3, ">" + dateTimeBegin.ToOADate(), Excel.XlAutoFilterOperator.xlAnd);
                ws.UsedRange.AutoFilter(4, "Плохо");
            } catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				RegistryMarksAddPivotTables(wb, xlApp, dataTable, dateTimeBegin);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			wb.Sheets[1].Name = "Негативные отзывы";
			wb.Sheets["Сводная таблица"].Activate();
			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static void RegistryMarksAddPivotTables(
            Excel.Workbook wb, Excel.Application xlApp, DataTable dataTable, DateTime dateTimeBegin) {
			SortedDictionary<string, ItemRegistryMark> marksSelectedPeriodByFilials = 
                new SortedDictionary<string, ItemRegistryMark>();
			SortedDictionary<string, SortedDictionary<string, ItemRegistryMark>> marksByWeeks 
                = new SortedDictionary<string, SortedDictionary<string, ItemRegistryMark>>();

            List<string> uniqueInnerKeys = new List<string>();

            foreach (DataRow dataRow in dataTable.Rows) {
				try {
                    DateTime createDate = DateTime.Parse(dataRow["createdate"].ToString());
                    
                    int weekNumber = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(
                        createDate, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday);
                    string dictByWeeksInnerKey = createDate.Year + " / " + weekNumber;

                    if (!uniqueInnerKeys.Contains(dictByWeeksInnerKey))
                        uniqueInnerKeys.Add(dictByWeeksInnerKey);

					string shortname = dataRow["SHORTNAME"].ToString();
					string department = dataRow["DEPARTMENT"].ToString();
                    string dictsOuterKey = shortname + " / " + department;

                    string mark = dataRow["MARK"].ToString();

                    if (!marksByWeeks.Keys.Contains(dictsOuterKey))
                        marksByWeeks.Add(dictsOuterKey, new SortedDictionary<string, ItemRegistryMark>());

                    if (!marksByWeeks[dictsOuterKey].Keys.Contains(dictByWeeksInnerKey))
                        marksByWeeks[dictsOuterKey].Add(dictByWeeksInnerKey, new ItemRegistryMark(shortname, department));

                    if (!marksSelectedPeriodByFilials.ContainsKey(dictsOuterKey))
						marksSelectedPeriodByFilials.Add(dictsOuterKey, new ItemRegistryMark(shortname, department));

					if (mark.Contains("Плохо")) {
                        marksByWeeks[dictsOuterKey][dictByWeeksInnerKey].MarkBad++;

                        if (createDate >= dateTimeBegin)
                            marksSelectedPeriodByFilials[dictsOuterKey].MarkBad++;
					} else if (mark.Contains("Средне")) {
                        marksByWeeks[dictsOuterKey][dictByWeeksInnerKey].MarkMedium++;

                        if (createDate >= dateTimeBegin)
                            marksSelectedPeriodByFilials[dictsOuterKey].MarkMedium++;
					} else if (mark.Contains("Хорошо")) {
                        marksByWeeks[dictsOuterKey][dictByWeeksInnerKey].MarkGood++;

                        if (createDate >= dateTimeBegin)
                            marksSelectedPeriodByFilials[dictsOuterKey].MarkGood++;
					} else if (mark.Contains("Дубль")) {
                        if (createDate >= dateTimeBegin)
                            marksSelectedPeriodByFilials[dictsOuterKey].MarkDuplicate++;
					} else
                        Logging.ToLog("Неизвестная оценка - " + mark);
                } catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			}

            uniqueInnerKeys.Sort();

            foreach (string innerKey in uniqueInnerKeys)
                foreach (string outerKey in marksByWeeks.Keys)
                    if (!marksByWeeks[outerKey].Keys.Contains(innerKey))
                        marksByWeeks[outerKey].Add(innerKey, new ItemRegistryMark("", ""));

            Excel.Worksheet ws = wb.Sheets["Сводная таблица"];
            ws.Activate();
            RegistryMarkDrawPivotTable(ws, marksSelectedPeriodByFilials);

            ws = wb.Sheets["График - кол-во"];
            ws.Activate();
            RegistryMarkDrawMarksByWeek(xlApp, ws, marksByWeeks, uniqueInnerKeys, RegistryMarkChartType.Total);

            //ws = wb.Sheets["График - %"];
            //ws.Activate();
            //RegistryMarkDrawMarksByWeek(xlApp, ws, marksByWeeks, uniqueInnerKeys, 1);

            //ws = wb.Sheets["График - KPI"];
            //ws.Activate();
            //RegistryMarkDrawMarksByWeek(xlApp, ws, marksByWeeks, uniqueInnerKeys, 2);

            Marshal.ReleaseComObject(ws);
            ws = null;
        }

        private enum RegistryMarkChartType {
            Total, Percentage, KPI
        }

        private static void RegistryMarkDrawPivotTable(
            Excel.Worksheet ws, SortedDictionary<string, ItemRegistryMark> marksSelectedPeriodByFilials) {
            int row = 2;
            int markBadTotal = 0;
            int markMediumTotal = 0;
            int markGoogTotal = 0;
            int markDuplicateTotal = 0;

            foreach (ItemRegistryMark item in marksSelectedPeriodByFilials.Values) {
                ws.Range["A" + row].Value = item.FilialName;
                ws.Range["B" + row].Value = item.Department;
                ws.Range["C" + row].Value = item.MarkBad;
                ws.Range["D" + row].Value = item.MarkMedium;
                ws.Range["E" + row].Value = item.MarkGood;
                ws.Range["F" + row].Value = item.MarkTotal;
                ws.Range["G" + row].Value = (item.MarkTotal > 0) ? (double)item.MarkBad / (double)item.MarkTotal : 0;
                ws.Range["H" + row].Value = (item.MarkTotal > 0) ? (double)item.MarkMedium / (double)item.MarkTotal : 0;
                ws.Range["I" + row].Value = (item.MarkTotal > 0) ? (double)item.MarkGood / (double)item.MarkTotal : 0;
                ws.Range["J" + row].Value = item.MarkDuplicate;

                markBadTotal += item.MarkBad;
                markMediumTotal += item.MarkMedium;
                markGoogTotal += item.MarkGood;
                markDuplicateTotal += item.MarkDuplicate;

                row++;
            }

            int totalMarks = markBadTotal + markMediumTotal + markGoogTotal;
            ws.Range["A" + row].Value = "Итого";
            ws.Range["C" + row].Value = markBadTotal;
            ws.Range["D" + row].Value = markMediumTotal;
            ws.Range["E" + row].Value = markGoogTotal;
            ws.Range["F" + row].Value = totalMarks;
            ws.Range["G" + row].Value = (totalMarks > 0) ? (double)markBadTotal / (double)totalMarks : 0;
            ws.Range["H" + row].Value = (totalMarks > 0) ? (double)markMediumTotal / (double)totalMarks : 0;
            ws.Range["I" + row].Value = (totalMarks > 0) ? (double)markGoogTotal / (double)totalMarks : 0;
            ws.Range["J" + row].Value = markDuplicateTotal;

            ws.Columns["G:I"].Style = "Percent";

            AddBoldBorder(ws.Range["A" + row + ":J" + row]);
            AddBoldBorder(ws.Range["A2:B" + row]);
            AddBoldBorder(ws.Range["C2:E" + row]);
            AddBoldBorder(ws.Range["F2:F" + row]);
            AddBoldBorder(ws.Range["G2:I" + row]);
            AddBoldBorder(ws.Range["J2:J" + row]);

            ws.Range["A" + row].Font.Bold = true;

            row += 2;
            ws.Range["A" + row].Value = "* попытки повторного голосования в течении 60 секунд";
            ws.Range["A" + row].Font.Italic = true;
            ws.Range["A1"].Select();
        }

        private static void RegistryMarkDrawMarksByWeek(
            Excel.Application xlApp, Excel.Worksheet ws, 
            SortedDictionary<string, SortedDictionary<string, ItemRegistryMark>> marksByWeeks, 
            List<string> uniqueInnerKeys,
            RegistryMarkChartType type) {
            string chartTitle;
            string hint;

            switch (type) {
                case RegistryMarkChartType.Total:
                    chartTitle = "Всего оценок - хронология";
                    hint = "Отображены все оценки 'Плохо' + 'Средне' + 'Хорошо'";
                    break;
                case RegistryMarkChartType.Percentage:
                    chartTitle = "Соотношение оценок хорошо и плохо";
                    hint = "Отображены только оценки 'Хорошо' и 'Плохо'";
                    break;
                case RegistryMarkChartType.KPI:
                    chartTitle = "KPI - хронология";
                    hint = "KPI рассчитывается по формуле: 'Средне' + 'Хорошо' / 'Всего'";
                    break;
                default:
                    Logging.ToLog("Неизвестный тип оценки - " + type);
                    return;
            }

            int row = 1;
            int column = 1;
            
            if (type != RegistryMarkChartType.Total)
                foreach (string innerKey in uniqueInnerKeys) {
                    ws.Cells[1, column].Value2 = innerKey;
                    column++;
                }

            row++;

            foreach (KeyValuePair<string, SortedDictionary<string, ItemRegistryMark>> pair in marksByWeeks) {
                string rowTitle = string.Empty;
                ws.Cells[row, 1].Value2 = pair.Key;
                column = 2;

                foreach (KeyValuePair<string, ItemRegistryMark> weekMarks in pair.Value) {
                    object value;
                    ItemRegistryMark mark = weekMarks.Value;

                    switch (type) {
                        case RegistryMarkChartType.Total:
                            string[] markDate = weekMarks.Key.Split('/');
                            string markYear = markDate[0].Trim(' ');
                            int weekNumber = Convert.ToInt32(markDate[1].Trim(' '));
                            string currentRowTitle = pair.Key + " " + markYear;

                            if (string.IsNullOrEmpty(rowTitle)) {
                                rowTitle = currentRowTitle;
                                ws.Cells[row, 1].Value2 = currentRowTitle;
                            } else {
                                if (!rowTitle.Equals(currentRowTitle)) {
                                    row++;
                                    rowTitle = currentRowTitle;
                                    ws.Cells[row, 1].Value2 = rowTitle;
                                }
                            }

                            column = weekNumber + 1;
                            value = mark.MarkBad + mark.MarkMedium + mark.MarkGood;
                            break;
                        case RegistryMarkChartType.Percentage:
                            value = weekMarks.Value.MarkGood;
                            break;
                        case RegistryMarkChartType.KPI:
                            if (mark.MarkTotal > 0)
                                value = ((double)mark.MarkTotal - (double)mark.MarkBad) / (double)mark.MarkTotal;
                            else
                                value = string.Empty;
                            break;
                        default:
                            continue;
                    }

                    ws.Cells[row, column].Value2 = value;

                    if (type == RegistryMarkChartType.Percentage || 
                        type == RegistryMarkChartType.KPI)
                        ws.Cells[row, column].NumberFormat = "0%";

                    column++;
                }

                row++;
            }

            column = 1;
            ws.Cells[row, column].Value2 = "ИТОГО";
            foreach (string innerKey in uniqueInnerKeys) {
                column++;
                if (type == RegistryMarkChartType.KPI) {
                    double marksPositive = 0;
                    double marksTotal = 0;
                    foreach (KeyValuePair<string, SortedDictionary<String, ItemRegistryMark>> pair in marksByWeeks) {
                        ItemRegistryMark mark = pair.Value[innerKey];
                        marksPositive += mark.MarkTotal - mark.MarkBad;
                        marksTotal += mark.MarkTotal;
                    }

                    if (marksTotal > 0)
                        ws.Cells[row, column].Value2 = marksPositive / marksTotal;
                    else
                        ws.Cells[row, column].Value2 = string.Empty;

                    ws.Cells[row, column].NumberFormat = "0%";
                } else 
                    ws.Cells[row, column].FormulaR1C1Local = "=СУММ(R[-" + (row - 2) + "]C:R[-1]C)";
            }

            Excel.Shape shape = xlApp.ActiveSheet.Shapes.AddChart2(234, Excel.XlChartType.xlLineMarkers, 10, 200, 1350, 370);
            shape.Select();
            xlApp.ActiveChart.SetSourceData(ws.UsedRange);
            xlApp.ActiveChart.ChartTitle.Text = chartTitle;

            for (int i = 1; i <= marksByWeeks.Keys.Count; i++)
                xlApp.ActiveChart.FullSeriesCollection(i).IsFiltered = true;

            ws.Cells[row + 2, 1].Value2 = hint;
            ws.Cells[row + 2, 1].Font.Italic = true;

            ws.Range["A1"].Select();

            Marshal.ReleaseComObject(shape);
            shape = null;
        }

		private class ItemRegistryMark {
			public string FilialName { get; private set; }
			public string Department { get; private set; }
			public int MarkBad { get; set; }
			public int MarkMedium { get; set; }
			public int MarkGood { get; set; }
			public int MarkDuplicate { get; set; }

            public int MarkTotal { get { return MarkBad + MarkMedium + MarkGood; } }

            public ItemRegistryMark(string filialName, string department) {
				FilialName = filialName;
				Department = department;
			}
		}



		public static string PerformUniqueServices(DataTable dataTableCurrent,
											 DataTable dataTableTotal,
											 string resultFilePrefix,
											 string templateName,
											 string period) {
			if (!GetTemplateFilePath(ref templateName))
				return string.Empty;

			string resultPath = GetResultFilePath(resultFilePrefix, templateName);

			try {
				File.Copy(templateName, resultPath);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				return string.Empty;
			}

			if (!OpenWorkbook(resultPath, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws))
				return string.Empty;

			Dictionary<string, int> serviceMap = new Dictionary<string, int> {
				{ "Имплантация (кол-во имплантов)", 7},
				{ "Протезирование: МК, включая коронки на имплантатах", 8},
				{ "Отбеливание ZOOM", 9},
				{ "Ортодонтия: кол-во начатых лечений (уник. обратившихся) в отчетный период", 10},
				{ "Направление на КЛКТ", 11},
				{ "ЭГДС под внутренней седацией", 12},
				{ "ФКС под внутренней седацией", 13},
				{ "Консультация диетолога", 14},
				{ "Закрытые направления КДЛ за наличный расчет (кол-во шт.)", 15}
			};

			Dictionary<string, string> filialMapCurrent = new Dictionary<string, string> {
				{ "МДМ", "G" },
				{ "СУЩ", "H" },
				{ "М-СРЕТ", "I" }
			};

			Dictionary<string, string> filialMapTotal = new Dictionary<string, string> {
				{ "МДМ", "K" },
				{ "СУЩ", "L" },
				{ "М-СРЕТ", "M" }
			};

			ParseAndWriteUniqueService(ws, dataTableCurrent, serviceMap, filialMapCurrent);
			ParseAndWriteUniqueService(ws, dataTableTotal, serviceMap, filialMapTotal);

			ws.Range["A1"].Value2 = ((string)ws.Range["A1"].Value2).Replace("@period", period);
			ws.Range["G5"].Value2 = ((string)ws.Range["G5"].Value2).Replace("@period", period);

			SaveAndCloseWorkbook(xlApp, wb, ws);
			return resultPath;
		}

		private static void ParseAndWriteUniqueService(Excel.Worksheet ws,
										  DataTable services,
										  Dictionary<string, int> serviceMap,
										  Dictionary<string, string> filialMap) {
			foreach (DataRow dataRow in services.Rows) {
				try {
					string filial = dataRow["SHORTNAME"].ToString().TrimStart(' ').TrimEnd(' ');
					string service = dataRow["SERVICE"].ToString().TrimStart(' ').TrimEnd(' ');
					int scount = Convert.ToInt32(dataRow["SCOUNT"].ToString().TrimStart(' ').TrimEnd(' '));

					if (!serviceMap.Keys.Contains(service) ||
						!filialMap.Keys.Contains(filial)) {
						Logging.ToLog("Не удалось найти ключи для пары: " + filial + "|" + service);
						continue;
					}

					ws.Range[filialMap[filial] + serviceMap[service]].Value2 = scount;
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			}
		}


		private static void AddBoldBorder(Excel.Range range) {
			try {
				//foreach (Excel.XlBordersIndex item in new Excel.XlBordersIndex[] {
				//	Excel.XlBordersIndex.xlDiagonalDown,
				//	Excel.XlBordersIndex.xlDiagonalUp,
				//	Excel.XlBordersIndex.xlInsideHorizontal,
				//	Excel.XlBordersIndex.xlInsideVertical}) 
				//	range.Borders[item].LineStyle = Excel.Constants.xlNone;

				foreach (Excel.XlBordersIndex item in new Excel.XlBordersIndex[] {
					Excel.XlBordersIndex.xlEdgeBottom,
					Excel.XlBordersIndex.xlEdgeLeft,
					Excel.XlBordersIndex.xlEdgeRight,
					Excel.XlBordersIndex.xlEdgeTop}) {
					range.Borders[item].LineStyle = Excel.XlLineStyle.xlContinuous;
					range.Borders[item].ColorIndex = 0;
					range.Borders[item].TintAndShade = 0;
					range.Borders[item].Weight = Excel.XlBorderWeight.xlMedium;
				}
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}
		}
	}
}
