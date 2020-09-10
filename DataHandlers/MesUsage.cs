using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	public class MesUsage : ExcelGeneral {
		private class ItemMESUsageTreatment {
			public string TREATDATE { get; set; } = string.Empty;
			public string FILIAL { get; set; } = string.Empty;
			public string DEPNAME { get; set; } = string.Empty;
			public string DOCNAME { get; set; } = string.Empty;
			public string HISTNUM { get; set; } = string.Empty;
			public string CLIENTNAME { get; set; } = string.Empty;
			public string MKBCODE { get; set; } = string.Empty;
			public string AGE { get; set; } = string.Empty;
			public Dictionary<string, int> DictMES { get; set; } = new Dictionary<string, int>(); //0 - Necessary, 1 - ByIndication, 2 - Additional, 3 - ExternalLPU
			public List<string> ListReferralsFromMes { get; set; } = new List<string>();
			public List<string> ListReferralsFromDoc { get; set; } = new List<string>();
			public Dictionary<string, ReferralDetails> DictAllReferrals { get; set; } = new Dictionary<string, ReferralDetails>();
			public string SERVICE_TYPE { get; set; } = string.Empty;
			public string PAYMENT_TYPE { get; set; } = string.Empty;
			public string AGNAME { get; set; } = string.Empty;
			public string AGNUM { get; set; } = string.Empty;
			public string PERIOD { get; set; } = string.Empty;

			public class ReferralDetails {
				public string Schid { get; set; } = string.Empty;
				public int IsCompleted { get; set; } = -1;  //0 - incompleted, 1 - completed
				public int RefType { get; set; } = -1; //2 - Lab
			}
		}


		private static Dictionary<string, ItemMESUsageTreatment> ParseMESUsageDataTableToTreatments(DataTable dataTable) {
			Dictionary<string, ItemMESUsageTreatment> treatments = new Dictionary<string, ItemMESUsageTreatment>();

			foreach (DataRow row in dataTable.Rows) {
				try {
					string treatcode = row["TREATCODE"].ToString();
					string mid = row["MID"].ToString();
					string listMES = row["LISTMES"].ToString();
					string listReferrals = row["LISTREFERRALS"].ToString();
					string listAllReferrals = row["LISTALLREFERRALS"].ToString();
					string[] arrayMES = new string[0];
					string[] arrayReferrals = new string[0];
					string[] arrayAllReferrals = new string[0];

					if (!string.IsNullOrEmpty(listMES))
						arrayMES = listMES.Split(';');

					if (!string.IsNullOrEmpty(listReferrals))
						arrayReferrals = listReferrals.Split(';');

					if (!string.IsNullOrEmpty(listAllReferrals))
						arrayAllReferrals = listAllReferrals.Split(';');

					if (treatments.ContainsKey(treatcode)) {
						foreach (KeyValuePair<string, int> pair in ParseMes(arrayMES))
							if (!treatments[treatcode].DictMES.ContainsKey(pair.Key))
								treatments[treatcode].DictMES.Add(pair.Key, pair.Value);

						if (string.IsNullOrEmpty(mid))
							treatments[treatcode].ListReferralsFromDoc.AddRange(arrayReferrals);
						else
							treatments[treatcode].ListReferralsFromMes.AddRange(arrayReferrals);
					} else {
						ItemMESUsageTreatment treatment = new ItemMESUsageTreatment() {
							TREATDATE = row["TREATDATE"].ToString(),
							CLIENTNAME = row["CLIENTNAME"].ToString(),
							HISTNUM = row["HISTNUM"].ToString(),
							DOCNAME = row["DOCNAME"].ToString(),
							FILIAL = row["FILIAL"].ToString(),
							DEPNAME = row["DEPNAME"].ToString(),
							MKBCODE = row["MKBCODE"].ToString(),
							AGE = row["AGE"].ToString(),
							AGNAME = row["AGNAME"].ToString(),
							AGNUM = row["AGNUM"].ToString(),
							SERVICE_TYPE = row["LISTALLSERVICES"].ToString().ToUpper().Contains("ПЕРВИЧНЫЙ") ? "Первичный" : "Повторный",
							PAYMENT_TYPE = "ДМС",
							PERIOD = row["PERIOD"].ToString()
						};

						if (!string.IsNullOrEmpty(row["GRNAME"].ToString()))
							if (row["NOGP"].ToString().Equals("1") && row["NODMS"].ToString().Equals("1"))
								treatment.PAYMENT_TYPE = "Нал";
							else continue;

						if (string.IsNullOrEmpty(mid))
							treatment.ListReferralsFromDoc.AddRange(arrayReferrals);
						else
							treatment.ListReferralsFromMes.AddRange(arrayReferrals);

						treatment.DictMES = ParseMes(arrayMES);
						treatment.DictAllReferrals = ParseAllReferrals(arrayAllReferrals);
						treatments.Add(treatcode, treatment);
					}
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			}

			return treatments;
		}

		private static Dictionary<string, ItemMESUsageTreatment.ReferralDetails> ParseAllReferrals(string[] valuesArray) {
			Dictionary<string, ItemMESUsageTreatment.ReferralDetails> keyValuePairs =
				new Dictionary<string, ItemMESUsageTreatment.ReferralDetails>();

			foreach (string item in valuesArray) {
				if (!item.Contains(":"))
					continue;

				try {
					string[] referral = item.Split(':');
					if (referral.Length < 3)
						continue;

					string referralCode = referral[0];

					if (keyValuePairs.ContainsKey(referralCode))
						continue;

					int.TryParse(referral[1], out int referralStatus);
					int.TryParse(referral[2], out int refType);
					ItemMESUsageTreatment.ReferralDetails referralDetails = new ItemMESUsageTreatment.ReferralDetails() {
						Schid = referralCode,
						IsCompleted = referralStatus,
						RefType = refType
					};

					keyValuePairs.Add(referralCode, referralDetails);
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			}
			
			return keyValuePairs;
		}

		private static Dictionary<string, int> ParseMes(string[] valuesArray) {
			Dictionary<string, int> keyValuePairs = new Dictionary<string, int>();

			foreach (string item in valuesArray) {
				if (!item.Contains(":"))
					continue;

				try {
					string[] referral = item.Split(':');
					string referralCode = referral[0];

					if (keyValuePairs.ContainsKey(referralCode))
						continue;

					int.TryParse(referral[1], out int referralStatus);
					keyValuePairs.Add(referralCode, referralStatus);
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			}

			return keyValuePairs;
		}


		//============================ MesUsage ============================
		public static string WriteMesUsageTreatmentsToExcel(DataTable dataTable, string resultFilePrefix, string templateFileName, bool isFullVersion = false) {
			if (!CreateNewIWorkbook(resultFilePrefix, templateFileName, out IWorkbook workbook, out ISheet sheet, out string resultFile, string.Empty))
				return string.Empty;

			Dictionary<string, ItemMESUsageTreatment> treatments = ParseMESUsageDataTableToTreatments(dataTable);

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
					treat.PERIOD,
					treat.TREATDATE, //Дата лечения
					treat.FILIAL, //Филиал
					treat.DEPNAME, //Подразделение
					treat.DOCNAME, //ФИО врача
					treat.HISTNUM, //Номер ИБ
					treat.CLIENTNAME, //ФИО пациента
					treat.AGE, //Возраст
					treat.MKBCODE, //Код МКБ
					necessaryServicesInMes, //Кол-во обязательных услуг согласно МЭС
					hasAtLeastOneReferralByMes, //Есть направление, созданное с использованием МЭС
					necessaryServiceReferralByMesInstrumental + necessaryServiceReferralByMesLaboratory, //Кол-во услуг в направлении с использованием МЭС
					hasAtLeastOneReferralSelfMade, //Есть направление, созданное самостоятельно
					necessaryServiceReferralSelfMadeInstrumental + necessaryServiceReferralSelfMadeLaboratory, //Кол-во услуг в направлении выставленных самостоятельно
					necessaryServiceInMesUsedPercent, //% Соответствия обязательных услуг МЭС (обязательные во всех направлениях) / всего обязательных в мэс
					necessaryServiceInMesUsedPercent == 1 ? 1 : 0, //Услуги из всех направлений соответсвуют обязательным услугам МЭС на 100%
					treat.SERVICE_TYPE, //Тип приема
					treat.PAYMENT_TYPE, //Тип оплаты приема
				};

				if (isFullVersion) {
					values = new List<object>() {
						treatment.Key, //Код лечения
						1, //Прием
						treat.PERIOD,
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
						necessaryServiceReferralCompletedByMesInstrumental, //Кол-во исполненных обязательных услуг в направлении МЭС (инструментальных)
						necessaryServiceReferralCompletedByMesLaboratory, //Кол-во исполненных обязательных услуг в направлении МЭС (лабораторных)
						hasAtLeastOneReferralSelfMade, //Есть направление, созданное самостоятельно
						necessaryServiceReferralSelfMadeInstrumental, //Кол-во обязательных услуг в направлении выставленных самостоятельно (инструментальных)
						necessaryServiceReferralSelfMadeLaboratory, //Кол-во обязательных услуг в направлении выставленных самостоятельно (лабораторных)
						necessaryServiceReferralCompletedSelfMadeInstrumental, //Кол-во исполненных обязательных услуг в самостоятельно созданных направлениях (инструментальных)
						necessaryServiceReferralCompletedSelfMadeLaboratory, //Кол-во исполненных обязательных услуг в самостоятельно созданных направлениях (лабораторных)
						servicesAllReferralsInstrumental, //Всего услуг во всех направлениях (иснтрументальных)
						servicesAllReferralsLaboratory, //Всего услуг во всех направлениях (лабораторных)
						completedServicesInReferrals, //Кол-во выполненных услуг во всех направлениях
						serviceInReferralOutsideMes, //Кол-во услуг в направлениях, не входящих в МЭС
						necessaryServiceInMesUsedPercent, //% Соответствия обязательных услуг МЭС (обязательные во всех направлениях) / всего обязательных в мэс
						necessaryServiceInMesUsedPercent == 1 ? 1 : 0, //Услуги из всех направлений соответсвуют обязательным услугам МЭС на 100%
						treat.SERVICE_TYPE, //Тип приема
						treat.PAYMENT_TYPE, //Тип оплаты приема
						treat.AGNAME, //Наименование организации
						treat.AGNUM //Номер договора
					};
				}

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

		public static bool Process(string resultFile, bool isFullVersion = false) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws))
				return false;

			try {
				ws.Activate();
				ws.Columns["D:D"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ";
				ws.Columns["Q:Q"].Select();
				xlApp.Selection.NumberFormat = "0%";
				ws.Range["A1"].Select();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			if (!isFullVersion) 
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
