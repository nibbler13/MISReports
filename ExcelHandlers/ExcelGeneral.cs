using Newtonsoft.Json;
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

namespace MISReports.ExcelHandlers {
	class ExcelGeneral {
		//============================ NPOI Excel ============================
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

		protected static bool GetTemplateFilePath(ref string templateFileName) {
			templateFileName = Path.Combine(Path.Combine(Program.AssemblyDirectory, "Templates\\"), templateFileName);

			if (!File.Exists(templateFileName)) {
				Logging.ToLog("Не удалось найти файл шаблона: " + templateFileName);
				return false;
			}

			return true;
		}

		public static string GetResultFilePath(string resultFilePrefix, string templateFileName = "", bool isPlainText = false) {
			string resultPath = Path.Combine(Program.AssemblyDirectory, "Results");
			if (!Directory.Exists(resultPath))
				Directory.CreateDirectory(resultPath);

			foreach (char item in Path.GetInvalidFileNameChars())
				resultFilePrefix = resultFilePrefix.Replace(item, '-');

			string fileEnding = ".xlsx";
			if (isPlainText)
				fileEnding = ".txt";

			string resultFile = Path.Combine(resultPath, resultFilePrefix + " " + DateTime.Now.ToString("yyyyMMdd_HHmmss") + fileEnding);

			if (isPlainText && !string.IsNullOrEmpty(templateFileName))
				File.Copy(templateFileName, resultFile, true);

			return resultFile;
		}

		protected static bool SaveAndCloseIWorkbook(IWorkbook workbook, string resultFile) {
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



		//============================ Interop Excel ============================
		protected static bool OpenWorkbook(string workbook, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws, string sheetName = "") {
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

		protected static void SaveAndCloseWorkbook(Excel.Application xlApp, Excel.Workbook wb, Excel.Worksheet ws) {
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

		public static bool CopyFormatting(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws))
				return false;

			try {
				int rowsUsed = ws.UsedRange.Rows.Count;
				string lastColumn = GetExcelColumnName(ws.UsedRange.Columns.Count);

				ws.Range["A2:" + lastColumn + "2"].Select();
				xlApp.Selection.Copy();
				ws.Range["A3:" + lastColumn + rowsUsed].Select();
				xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
				ws.Rows["2:" + rowsUsed].Select();
				xlApp.Selection.RowHeight = 15;

				ws.Range["A1"].Select();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static string GetExcelColumnName(int columnNumber) {
			int dividend = columnNumber;
			string columnName = String.Empty;
			int modulo;

			while (dividend > 0) {
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}

			return columnName;
		}



		public static string WriteDataTableToExcel(DataTable dataTable, string resultFilePrefix, string templateFileName,
			string sheetName = "", bool createNew = true, string workloadFilial = "", ReportsInfo.Type? type = null) {
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
			bool telemedicineOnlyIngosstrakh = false;

			if (type.HasValue) {
				if (type.Value == ReportsInfo.Type.PriceListToSite || type.Value == ReportsInfo.Type.FssInfo)
					rowNumber = 2;

				if (type.Value == ReportsInfo.Type.TelemedicineOnlyIngosstrakh)
					telemedicineOnlyIngosstrakh = true;
			}

            List<string> valuesToClearFssInfo = new List<string> { "0", "-1", ";", ";;", "01.01.0001 0:00:00" };

            foreach (DataRow dataRow in dataTable.Rows) {
				if (!string.IsNullOrEmpty(workloadFilial) && !workloadFilial.Equals("_Общий")) {
					string currentRowFilial = dataRow[3].ToString();

					if (!currentRowFilial.StartsWith(workloadFilial))
						continue;
				}

				IRow row = null;
				try { row = sheet.GetRow(rowNumber); } catch (Exception) { }

				if (row == null)
					row = sheet.CreateRow(rowNumber);

				if (telemedicineOnlyIngosstrakh)
					try {
						string paymentType = dataRow["JNAME"].ToString();
						if (!paymentType.ToLower().Contains("ингосстрах"))
							continue;
					} catch (Exception) { }
				
				foreach (DataColumn column in dataTable.Columns) {
					ICell cell = null;
					try { cell = row.GetCell(columnNumber); } catch (Exception) { }

					if (cell == null)
						cell = row.CreateCell(columnNumber);

					string value = dataRow[column].ToString();

                    if (type.HasValue && type.Value == ReportsInfo.Type.PriceListToSite && (columnNumber == 4 || columnNumber == 7)) {
                        cell.SetCellValue(value);
                    } else if (type.HasValue && type.Value == ReportsInfo.Type.FssInfo && valuesToClearFssInfo.Contains(value)) { 
                        cell.SetCellValue(string.Empty);
                    } else {
                        if (double.TryParse(value, out double result)) {
                            cell.SetCellValue(result);
                        } else if (DateTime.TryParse(value, out DateTime date)) {
                            cell.SetCellValue(date);
                        } else {
                            cell.SetCellValue(value);
                        }
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

		protected static void AddBoldBorder(Excel.Range range) {
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

		protected static void AddInteriorColor(Excel.Range range, Excel.XlThemeColor xlThemeColor) {
			range.Interior.Pattern = Excel.Constants.xlSolid;
			range.Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
			range.Interior.ThemeColor = xlThemeColor;
			range.Interior.TintAndShade = 0.799981688894314;
			range.Interior.PatternTintAndShade = 0;
		}



		//============================ OleDB Excel ============================
		public static DataTable ReadExcelFile(string fileName, string sheetName) {
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


        //=============================== MISC ================================
        public static string ColumnIndexToColumnLetter(int colIndex) {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0) {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }

            return colLetter;
        }



        public static string WriteDataTableToTextFile(DataTable dataTable, string resultFilePrefix = "", string templateFileName = "", bool saveAsJson = false) {
            string resultFile = string.Empty;

            try {
                if (saveAsJson) {
                    string json = JsonConvert.SerializeObject(dataTable, Formatting.Indented);
                    string filePath = GetResultFilePath(resultFilePrefix, "", true);

                    File.WriteAllText(filePath, json);
                    resultFile = filePath;
                } else {
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
                }
            } catch (Exception e) {
                Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
            }

            return resultFile;
        }
    }
}
