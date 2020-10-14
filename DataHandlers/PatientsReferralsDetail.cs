using MISReports.ExcelHandlers;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.DataHandlers {
	class PatientsReferralsDetail : ExcelGeneral {
		static List<ItemTreatment> ParseDataTable(DataTable dataTable) {
			List<ItemTreatment> treatments = new List<ItemTreatment>();

			foreach (DataRow row in dataTable.Rows) {
				string treatdate = row["TREATDATE"].ToString();
				string treatcode = row["TREATCODE"].ToString();
				string filial = row["FILIAL"].ToString();
				string client = row["CLIENT"].ToString();
				string pcode = row["PCODE"].ToString();
				string histnum = row["HISTNUM"].ToString();
				string phone = row["PHONE"].ToString();
				string depname = row["DEPNAME"].ToString();
				string dcode = row["DCODE"].ToString();
				string doc = row["DOC"].ToString();
				string services = row["SERVICES"].ToString();
				string referralsAll = row["REFERRALS_ALL"].ToString();
				string referralsScheduled = row["REFERRALS_SCHEDULED"].ToString();
				string referralsCompleted = row["REFERRALS_COMPLETED"].ToString();

				string[] phoneSplitted = phone.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
				phone = string.Join(", ", phoneSplitted);

				ItemTreatment itemTreatment = new ItemTreatment {
					Treatdate = treatdate,
					Treatcode = treatcode,
					Filial = filial,
					Client = client,
					Pcode = pcode,
					Histnum = histnum,
					Phone = phone,
					TreatType = services.ToLower().Contains("первичный") ? "Первичный" : (services.ToLower().Contains("повторный") ? "Повторный" : ""),
					Depname = depname,
					Dcode = dcode,
					Doc = doc
				};

				string[] servicesSplitted = services.Split(new char[] { '^' }, StringSplitOptions.RemoveEmptyEntries);
				foreach (string service in servicesSplitted)
					if (ParseService(service, row, out ItemService itemService))
						itemTreatment.Services.Add(itemService);

				string[] referralsAllSplitted = referralsAll.Split(new char[] { '#' }, StringSplitOptions.RemoveEmptyEntries);
				foreach (string referall in referralsAllSplitted) {
					string[] referralSplitted = referall.Split('^');
					if (referralSplitted.Length != 3) {
						Console.WriteLine("referralsAll: " + referralsAll);
						Logging.ToLog("Неправильная длина объекта направление, должно быть 3 части, сейчас: " + referralSplitted.Length);
						continue;
					}
					ItemReferral itemReferral = new ItemReferral {
						Refid = long.Parse(referralSplitted[0]),
						RefType = referralSplitted[1]
					};

					string[] refServicesSplitted = referralSplitted[2].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
					foreach (string service in refServicesSplitted)
						if (ParseService(service, row, out ItemService itemService))
							itemReferral.Services.Add(itemService);

					itemTreatment.Referrals.Add(itemReferral.Refid, itemReferral);
				}

				string[] referralsScheduledSplitted = referralsScheduled.Split(new char[] { '^' }, StringSplitOptions.RemoveEmptyEntries);
				foreach (string referralScheduled in referralsScheduledSplitted) {
					string[] scheduleSplitted = referralScheduled.Split('$');
					if (scheduleSplitted.Length != 8) {
						Console.WriteLine("referralsScheduled: " + referralsScheduled);
						Logging.ToLog("Неправильная длина объекта запись, должно быть 8 частей, сейчас: " + scheduleSplitted.Length);
						continue;
					}

					try {
						long refid = long.Parse(scheduleSplitted[0]);

						ItemSchedule itemSchedule = new ItemSchedule {
							Schedid = long.Parse(scheduleSplitted[1]),
							Workdate = DateTime.Parse(scheduleSplitted[2]),
							Doc = scheduleSplitted[3],
							DocPost = scheduleSplitted[4],
							AuthorFilial = scheduleSplitted[5],
							AuthorName = scheduleSplitted[6],
							AuthorPost = scheduleSplitted[7]
						};

						if (itemTreatment.Referrals.ContainsKey(refid))
							itemTreatment.Referrals[refid].Schedule = itemSchedule;
					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					}
				}

				string[] referralsCompletedSplitted = referralsCompleted.Split(new char[] { '#' }, StringSplitOptions.RemoveEmptyEntries);
				foreach (string referralCompleted in referralsCompletedSplitted) {
					string[] treatSplitted = referralCompleted.Split('^');
					if (treatSplitted.Length != 4) {
						Console.WriteLine("referralsCompleted: " + referralsCompleted);
						Logging.ToLog("Неправильная длина объекта прием, должно быть 4 части, сейчас: " + treatSplitted.Length);
						continue;
					}

					try {
						long refid = long.Parse(treatSplitted[0]);

						ItemTreat itemTreat = new ItemTreat {
							Treatcode = long.Parse(treatSplitted[1]),
							Treatdate = DateTime.Parse(treatSplitted[2])
						};

						string treatServices = treatSplitted[3];
						string[] treatServicesSplitted = treatServices.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
						foreach (string treatService in treatServicesSplitted)
							if (ParseService(treatService, row, out ItemService itemService))
								itemTreat.Services.Add(itemService);

						if (itemTreatment.Referrals.ContainsKey(refid))
							itemTreatment.Referrals[refid].Treat = itemTreat;
					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					}

				}

				treatments.Add(itemTreatment);
			}

			return treatments;
		}

		class ItemTreatment {
			public string Treatdate { get; set; }
			public string Treatcode { get; set; }
			public string Filial { get; set; }
			public string Client { get; set; }
			public string Pcode { get; set; }
			public string Histnum { get; set; }
			public string Phone { get; set; }
			public string TreatType { get; set; }
			public string Depname { get; set; }
			public string Dcode { get; set; }
			public string Doc { get; set; }
			public List<ItemService> Services { get; } = new List<ItemService>();
			public Dictionary<long, ItemReferral> Referrals { get; } = new Dictionary<long, ItemReferral>();
		}

		class ItemService {
			public string Kodoper { get; set; }
			public string Name { get; set; }
			public int Count { get; set; }
			public double Amount { get; set; }
			public bool IsUsed { get; set; } = false;
		}

		class ItemReferral {
			public long Refid { get; set; }
			public string RefType { get; set; }
			public List<ItemService> Services { get; } = new List<ItemService>();
			public ItemSchedule Schedule { get; set; }
			public ItemTreat Treat { get; set; }
		}

		class ItemSchedule {
			public long Schedid { get; set; }
			public DateTime Workdate { get; set; }
			public string Doc { get; set; }
			public string DocPost { get; set; }
			public string AuthorFilial { get; set; }
			public string AuthorName { get; set; }
			public string AuthorPost { get; set; }
		}

		class ItemTreat {
			public long Treatcode { get; set; }
			public DateTime Treatdate { get; set; }
			public List<ItemService> Services { get; } = new List<ItemService>();
		}

		static bool ParseService(string service, DataRow item, out ItemService itemService) {
			itemService = null;
			string[] serviceSplitted = service.Split('$');
			if (serviceSplitted.Length != 4) {
				Console.WriteLine("service: " + service);
				Logging.ToLog("Неправильная длина объекта услуга, должно быть 4 части, сейчас: " + serviceSplitted.Length);
				return false;
			}

			try {
				itemService = new ItemService {
					Kodoper = serviceSplitted[0],
					Name = serviceSplitted[1],
					Count = int.Parse(serviceSplitted[2]),
					Amount = double.Parse(serviceSplitted[3].Replace(".", ","))
				};

				return true;
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			return false;
		}

		public static string WriteToExcel(DataTable dataTable, string resultFilePrefix, string templateFileName) {
			if (!CreateNewIWorkbook(resultFilePrefix, templateFileName, out IWorkbook workbook, out ISheet sheet, out string resultFile, string.Empty))
				return string.Empty;

			List<ItemTreatment> treatments = ParseDataTable(dataTable);

			int rowNumber = 1;
			int columnNumber = 0;

			foreach (ItemTreatment treatment in treatments) {
				List<object> valuesTreatment = new List<object>() {
					treatment.Treatdate,
					treatment.Filial,
					treatment.Treatcode,
					treatment.Client,
					treatment.Pcode,
					treatment.Histnum,
					treatment.Phone,
					treatment.TreatType,
					treatment.Depname,
					treatment.Dcode,
					treatment.Doc
				};

				int serviceCountResult = 0;
				int serviceQuantityResult = 0;
				double serviceAmountResult = 0;

				int serviceRow = rowNumber;
				for (int treSer = 0; treSer < treatment.Services.Count; treSer++) {
					ItemService service = treatment.Services[treSer];

					serviceCountResult += service.Count;
					serviceQuantityResult++;
					serviceAmountResult += service.Amount;

					List<object> valuesService = new List<object>() {
						service.Name,
						service.Kodoper,
						service.Count,
						service.Amount
					};

					List<object> valuesToWrite = new List<object>();
					valuesToWrite.AddRange(valuesTreatment);
					valuesToWrite.AddRange(valuesService);

					if (treSer > 0) {
						valuesToWrite[2] = null;
						valuesToWrite[4] = null;
						valuesToWrite[9] = null;
					}

					WriteOutRow(valuesToWrite, sheet, serviceRow, columnNumber);
					serviceRow++;
				}

				int referralQuantityResult = 0;
				int referralServiceQuantityResult = 0;
				int referralServiceCountResult = 0;
				double referralServiceAmountResult = 0;

				int referralScheduleQuantityResult = 0;
				int referralTreatQuantityResult = 0;
				int referralTreatServiceQuantityResult = 0;
				int referralTreatServiceCountResult = 0;
				double referralTreatServiceAmountResult = 0;

				int referralRow = rowNumber;
				for (int treRef = 0; treRef < treatment.Referrals.Count; treRef++) {
					ItemReferral referral = treatment.Referrals.ElementAt(treRef).Value;
					referralQuantityResult++;
					List<long> treatcodesByReferrals = new List<long>();

					for (int refSer = 0; refSer < referral.Services.Count; refSer++) {
						ItemService service = referral.Services[refSer];
						referralServiceQuantityResult++;
						referralServiceCountResult += service.Count;
						referralServiceAmountResult += service.Amount;

						List<object> valuesToWrite = new List<object>();
						valuesToWrite.AddRange(valuesTreatment);
						valuesToWrite.AddRange(new List<object> {
							null,
							null,
							null,
							null
						});

						List<object> valuesService = new List<object> {
							referral.RefType,
							refSer == 0 ? referral.Refid.ToString() : null,
							service.Name,
							service.Kodoper,
							service.Count,
							service.Amount
						};

						if (referral.Schedule is null)
							valuesService.AddRange(new List<object> {
								"Отсутствует",
								null,
								null,
								null,
								null,
								null,
								null,
								null
							});
						else {
							referralScheduleQuantityResult++;

							valuesService.AddRange(new List<object> {
								referral.Schedule.Workdate,
								refSer == 0 ? referral.Schedule.Schedid.ToString() : null,
								refSer == 0 ? treatment.Pcode.ToString() : null,
								referral.Schedule.Doc,
								referral.Schedule.DocPost,
								referral.Schedule.AuthorFilial,
								referral.Schedule.AuthorName,
								referral.Schedule.AuthorPost
							});
						}

						if (referral.Treat is null)
							valuesService.AddRange(new List<object> {
								"Нет",
								null,
								null,
								null,
								null,
								null,
								null,
								null
							});
						else {

							if (treatcodesByReferrals.Contains( referral.Treat.Treatcode)) {
								valuesService.AddRange(new List<object> {
									"Да",
									referral.Treat.Treatdate,
									null,
									null
								});
							} else {
								valuesService.AddRange(new List<object> {
									"Да",
									referral.Treat.Treatdate,
									referral.Treat.Treatcode,
									treatment.Pcode
								});

								referralTreatQuantityResult++;
								treatcodesByReferrals.Add(referral.Treat.Treatcode);
							}

							foreach (ItemService treatService in referral.Treat.Services) {
								if (treatService.IsUsed)
									continue;

								if (!treatService.Name.Equals(service.Name))
									continue;

								referralTreatServiceQuantityResult++;
								referralTreatServiceCountResult += treatService.Count;
								referralTreatServiceAmountResult += treatService.Amount;

								valuesService.AddRange(new List<object> {
									treatService.Name,
									treatService.Kodoper,
									treatService.Count,
									treatService.Amount
								});

								valuesService.Add(
									((double)treatService.Amount / (double)service.Amount * 100.0d)
									.ToString("N2", CultureInfo.CurrentCulture));

								treatService.IsUsed = true;
								break;
							}
						}

						valuesToWrite.AddRange(valuesService);

						if (treRef > 0) {
							valuesToWrite[2] = null;
							valuesToWrite[4] = null;
							valuesToWrite[9] = null;
						}

						WriteOutRow(valuesToWrite, sheet, referralRow, columnNumber);
						referralRow++;
					}
				}

				columnNumber = 0;

				if (serviceRow != rowNumber || referralRow != rowNumber)
					rowNumber = serviceRow > referralRow ? serviceRow : referralRow;
				else {
					WriteOutRow(valuesTreatment, sheet, rowNumber, columnNumber);
					rowNumber++;
				}

				List<object> valuesToWriteResult = new List<object>();
				valuesToWriteResult.AddRange(valuesTreatment);

				valuesToWriteResult.Add("Итого по услугам");
				valuesToWriteResult.Add(serviceQuantityResult);
				valuesToWriteResult.Add(serviceCountResult);
				valuesToWriteResult.Add(serviceAmountResult);

				valuesToWriteResult.Add("Итого по направлениям");
				valuesToWriteResult.Add(referralQuantityResult);
				valuesToWriteResult.Add(null);
				valuesToWriteResult.Add(referralServiceQuantityResult);
				valuesToWriteResult.Add(referralServiceCountResult);
				valuesToWriteResult.Add(referralServiceAmountResult);

				valuesToWriteResult.Add("Итого по назначениям");
				valuesToWriteResult.Add(referralScheduleQuantityResult);

				valuesToWriteResult.AddRange(new List<object> {
					null,
					null,
					null,
					null,
					null,
					null
				});

				string serviceResult = "Нет услуг в направлениях";
				if (referralServiceCountResult > 0)
					serviceResult = ((double)referralTreatServiceCountResult / (double)referralServiceCountResult * 100.0)
						.ToString("N2", CultureInfo.CurrentCulture) + "% - исполнения (по услугам)";
				valuesToWriteResult.Add(serviceResult);

				valuesToWriteResult.Add("Итого по приемам");
				valuesToWriteResult.Add(referralTreatQuantityResult);
				valuesToWriteResult.Add(null);
				valuesToWriteResult.Add(null);
				valuesToWriteResult.Add(referralTreatServiceQuantityResult);
				valuesToWriteResult.Add(referralTreatServiceCountResult);
				valuesToWriteResult.Add(referralTreatServiceAmountResult);

				if (referralServiceAmountResult > 0)
					valuesToWriteResult.Add(
						((double)referralTreatServiceAmountResult / (double)referralServiceAmountResult * 100.0d)
						.ToString("N2", CultureInfo.CurrentCulture));
				else
					valuesToWriteResult.Add("Нет направлений");

				valuesToWriteResult[2] = null;
				valuesToWriteResult[4] = null;
				valuesToWriteResult[9] = null;

				WriteOutRow(valuesToWriteResult, sheet, rowNumber, columnNumber);
				columnNumber = 0;
				rowNumber++;
			}

			if (!SaveAndCloseIWorkbook(workbook, resultFile))
				return string.Empty;

			return resultFile;
		}


		private static void WriteOutRow(List<object> values, ISheet sheet, int rowNumber, int columnNumber) {
			IRow row = null;
			try { row = sheet.GetRow(rowNumber); } catch (Exception) { }

			if (row == null)
				row = sheet.CreateRow(rowNumber);

			foreach (object value in values) {
				if (value != null) {
					ICell cell = null;
					try { cell = row.GetCell(columnNumber); } catch (Exception) { }

					if (cell == null)
						cell = row.CreateCell(columnNumber);

					if (double.TryParse(value.ToString(), out double result))
						cell.SetCellValue(result);
					else if (DateTime.TryParse(value.ToString(), out DateTime date))
						cell.SetCellValue(date);
					else
						cell.SetCellValue(value.ToString());
				}

				columnNumber++;
			}
		}

		public static bool Process(string resultFile) {
			if (!CopyFormatting(resultFile))
				return false;

			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws))
				return false;

			try {
				int usedRows = ws.UsedRange.Rows.Count;
				ws.Range["A1"].Select();
				bool isDark = true;

				for (int i = 2; i < usedRows; i++) {
					object tr = ws.Range["C" + i].Value2;
					string treatcode = tr is null ? string.Empty : tr.ToString();

					for (int y = i + 1; y <= usedRows; y++) {
						object trNext = ws.Range["C" + y].Value2;
						string treatcodeNext = trNext is null ? string.Empty : trNext.ToString();
						if ((string.IsNullOrEmpty(treatcodeNext) || treatcodeNext.Equals(treatcode)) && y != usedRows)
							continue;

						int rowEnd = y - 1;
						if (y == usedRows)
							rowEnd = usedRows;

						SetColorForRange(ws.Range["A" + rowEnd + ":AR" + rowEnd], isDark);
						i = rowEnd;
						break;
					}
				}

				ws.Range["AM2:AR2"].Select();
				xlApp.Selection.AutoFill(ws.Range["AM2:AR" + usedRows], Excel.XlAutoFillType.xlFillValues);
				ws.Range["AM3:AR3"].Select();
				xlApp.Selection.AutoFill(ws.Range["AM2:AR3"], Excel.XlAutoFillType.xlFillValues);
				ws.Range["A1"].Select();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				AddPivotTableScheduleResult(wb, ws, xlApp);
				AddPivotTableReferralResult(wb, ws, xlApp);
				AddPivotTableReferralUsage(wb, ws, xlApp);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static void SetColorForRange(Excel.Range rangeToColorize, bool isDark) {
			double tintAndShade = -1;
			if (isDark)
				tintAndShade = 0.8;
			else 
				tintAndShade = 0;

			if (tintAndShade != 0) {
				rangeToColorize.Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
				rangeToColorize.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4;
				rangeToColorize.Interior.TintAndShade = tintAndShade;
				rangeToColorize.Interior.PatternTintAndShade = 0;
			} else {
				rangeToColorize.Interior.Pattern = Excel.Constants.xlNone;
				rangeToColorize.Interior.TintAndShade = 0;
				rangeToColorize.Interior.PatternTintAndShade = 0;
			}
		}


		private static void AddPivotTableScheduleResult(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			string pivotTableName = @"SchedResult";
			Excel.Worksheet wsPivote = wb.Sheets["Кол-во записей, итог по сумме"];

			int rowsUsed = ws.UsedRange.Rows.Count;

			wsPivote.Activate();

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, "Данные!R1C1:R" + rowsUsed + "C44", 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			//pivotTable.PivotFields("Дата").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			//pivotTable.PivotFields("Дата").Position = 1;

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 1;

			pivotTable.PivotFields("Запись, Должность").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Запись, Должность").Position = 2;

			pivotTable.PivotFields("Запись, Автор").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Запись, Автор").Position = 3;

			pivotTable.PivotFields("Услуга").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
			pivotTable.PivotFields("Услуга").Position = 1;
			pivotTable.PivotFields("Услуга").CurrentPage = "All";
			pivotTable.PivotFields("Услуга").PivotItems("Итого по услугам").Visible = false;
			pivotTable.PivotFields("Услуга").EnableMultiplePageItems = true;

			pivotTable.AddDataField(pivotTable.PivotFields("Кол-во"), "Планируемое кол-во услуг", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Планируемое кол-во услуг").NumberFormat = "# ##0";

			pivotTable.AddDataField(pivotTable.PivotFields("Планируемая стоимость, всего"), "Планируемая стоимость (всего)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Планируемая стоимость (всего)").NumberFormat = "# ##0,00 ?";

			pivotTable.AddDataField(pivotTable.PivotFields("Schedid"), "Кол-во записей (по направлениям)", Excel.XlConsolidationFunction.xlCount);
			pivotTable.PivotFields("Кол-во записей (по направлениям)").NumberFormat = "# ##0";

			pivotTable.AddDataField(pivotTable.PivotFields("Стоимость, всего2"), "Сумма оказанных услуг (по направлениям)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Сумма оказанных услуг (по направлениям)").NumberFormat = "# ##0,00 ?";

			pivotTable.PivotFields("Запись, Должность").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Запись, Должность").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Филиал").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Филиал").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Дата").Subtotals = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Дата").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Запись, Должность").ShowDetail = false;
			pivotTable.PivotFields("Филиал").ShowDetail = false;
			//pivotTable.PivotFields("Дата").ShowDetail = false;

			wb.ShowPivotTableFieldList = false;
			//pivotTable.DisplayFieldCaptions = false;

			wsPivote.Range["A1"].Select();
		}


		private static void AddPivotTableReferralResult(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			string pivotTableName = @"RefResult";
			Excel.Worksheet wsPivote = wb.Sheets["Реализация направлений"];

			int rowsUsed = ws.UsedRange.Rows.Count;

			wsPivote.Activate();

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, "Данные!R1C1:R" + rowsUsed + "C44", 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			//pivotTable.PivotFields("Дата").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			//pivotTable.PivotFields("Дата").Position = 1;

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 1;

			pivotTable.AddDataField(pivotTable.PivotFields("Refid"), "Кол-во направлений (всего)", Excel.XlConsolidationFunction.xlCount);
			pivotTable.PivotFields("Кол-во направлений (всего)").NumberFormat = "# ##0";

			pivotTable.AddDataField(pivotTable.PivotFields("Schedid"), "Кол-во записей (по направлениям)", Excel.XlConsolidationFunction.xlCount);
			pivotTable.PivotFields("Кол-во записей (по направлениям)").NumberFormat = "# ##0";

			pivotTable.CalculatedFields().Add("% Записей_", "='Кол-во записей'/'Кол-во приемов, исходных'", true);
			pivotTable.PivotFields("% Записей_").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю % Записей_").NumberFormat = "0,00%";
			pivotTable.PivotFields("Сумма по полю % Записей_").Caption = "% Записей";

			pivotTable.AddDataField(pivotTable.PivotFields("Treatcode2"), "Кол-во приемов (по направлениям)", Excel.XlConsolidationFunction.xlCount);
			pivotTable.PivotFields("Кол-во приемов (по направлениям)").NumberFormat = "# ##0";

			pivotTable.CalculatedFields().Add("% Приемов (по направлениям)_", "='Кол-во приемов, по направлениям'/'Кол-во приемов, исходных'", true);
			pivotTable.PivotFields("% Приемов (по направлениям)_").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю % Приемов (по направлениям)_").NumberFormat = "0,00%";
			pivotTable.PivotFields("Сумма по полю % Приемов (по направлениям)_").Caption = "% Приемов (по направлениям)";

			pivotTable.AddDataField(pivotTable.PivotFields("Стоимость, всего2"), "Сумма оказанных услуг (по направлениям)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Сумма оказанных услуг (по направлениям)").NumberFormat = "# ##0,00 ?";

			pivotTable.AddDataField(pivotTable.PivotFields("Уникальность пациента, исходный прием"), "Кол-во уникальных пациентов (всего)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Кол-во уникальных пациентов (всего)").NumberFormat = "# ##0 ?";

			pivotTable.AddDataField(pivotTable.PivotFields("Уникальность пациента, запись"), "Кол-во уникальных пациентов (запись)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Кол-во уникальных пациентов (запись)").NumberFormat = "# ##0";

			pivotTable.CalculatedFields().Add("% Записанных пациентов_", "='Уникальность пациента, запись'/'Уникальность пациента, исходный прием'", true);
			pivotTable.PivotFields("% Записанных пациентов_").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю % Записанных пациентов_").NumberFormat = "0,00%";
			pivotTable.PivotFields("Сумма по полю % Записанных пациентов_").Caption = "% Записанных пациентов";

			pivotTable.AddDataField(pivotTable.PivotFields("Уникальность пациента, прием по направлению"), "Кол-во уникальных пациентов (прием по направлениям)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Кол-во уникальных пациентов (прием по направлениям)").NumberFormat = "# ##0";

			pivotTable.CalculatedFields().Add("% Принятых по направлению пациентов_", "='Уникальность пациента, прием по направлению'/'Уникальность пациента, исходный прием'", true);
			pivotTable.PivotFields("% Принятых по направлению пациентов_").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю % Принятых по направлению пациентов_").NumberFormat = "0,00%";
			pivotTable.PivotFields("Сумма по полю % Принятых по направлению пациентов_").Caption = "% Принятых по направлению пациентов";


			//ActiveSheet.PivotTables("RefResult").CalculatedFields.Add "% Записей", "='Кол-во записей'/'Кол-во приемов, исходных'", True
			//ActiveSheet.PivotTables("RefResult").PivotFields("% Записей").Orientation = xlDataField

			//ActiveSheet.PivotTables("RefResult").CalculatedFields.Add "% Приемов (по направлениям)", "='Кол-во приемов, по направлениям'/'Кол-во приемов, исходных'", True
			//ActiveSheet.PivotTables("RefResult").PivotFields("% Приемов (по направлениям)").Orientation = xlDataField

			//ActiveSheet.PivotTables("RefResult").CalculatedFields.Add "% Записанных пациентов", "='Уникальность пациента, запись'/'Уникальность пациента, исходный прием'", True
			//ActiveSheet.PivotTables("RefResult").PivotFields("% Записанных пациентов").Orientation = xlDataField

			//ActiveSheet.PivotTables("RefResult").CalculatedFields.Add "% Принятых по направлению пациентов", "='Уникальность пациента, прием по направлению'/'Уникальность пациента, исходный прием'", True
			//ActiveSheet.PivotTables("RefResult").PivotFields("% Принятых по направлению пациентов").Orientation = xlDataField

			pivotTable.PivotFields("Услуга").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
			pivotTable.PivotFields("Услуга").Position = 1;
			pivotTable.PivotFields("Услуга").CurrentPage = "All";
			pivotTable.PivotFields("Услуга").PivotItems("Итого по услугам").Visible = false;
			pivotTable.PivotFields("Услуга").EnableMultiplePageItems = true;

			//pivotTable.AddDataField(pivotTable.PivotFields("№ ИБ"), "Кол-во записей", Excel.XlConsolidationFunction.xlCount);

			//pivotTable.PivotFields("Дата").ShowDetail = false;

			wb.ShowPivotTableFieldList = false;
			//pivotTable.DisplayFieldCaptions = false;

			pivotTable.HasAutoFormat = false;
			wsPivote.Columns["B:L"].ColumnWidth = 15;
			wsPivote.Range["B3:L3"].VerticalAlignment = Excel.Constants.xlTop;
			wsPivote.Range["B3:L3"].WrapText = true;

			wsPivote.Range["A1"].Select();
		}


		private static void AddPivotTableReferralUsage(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			string pivotTableName = @"RefUsage";
			Excel.Worksheet wsPivote = wb.Sheets["Использование направлений"];

			int rowsUsed = ws.UsedRange.Rows.Count;

			wsPivote.Activate();

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, "Данные!R1C1:R" + rowsUsed + "C44", 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			//pivotTable.PivotFields("Дата").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			//pivotTable.PivotFields("Дата").Position = 1;

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 1;

			pivotTable.PivotFields("Отделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Отделение").Position = 2;

			pivotTable.PivotFields("ФИО Сотрудника").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ФИО Сотрудника").Position = 3;


			pivotTable.PivotFields("Услуга").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
			pivotTable.PivotFields("Услуга").Position = 1;
			pivotTable.PivotFields("Услуга").CurrentPage = "Итого по услугам";
			//pivotTable.PivotFields("Услуга").PivotItems("Итого по услугам").Visible = true;
			//pivotTable.PivotFields("Услуга").EnableMultiplePageItems = false;


			pivotTable.AddDataField(pivotTable.PivotFields("ФИО Пациента"), "Кол-во приемов", Excel.XlConsolidationFunction.xlCount);
			pivotTable.PivotFields("Кол-во приемов").NumberFormat = "# ##0";

			pivotTable.AddDataField(pivotTable.PivotFields("Стоимость, всего"), "Сумма оказанных услуг", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Сумма оказанных услуг").NumberFormat = "# ##0,00 ?";

			pivotTable.CalculatedFields().Add("СрЧек", "='Стоимость, всего'/'Кол-во приемов, исходных'", true);
			pivotTable.PivotFields("СрЧек").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю СрЧек").NumberFormat = "# ##0,00 ?";
			pivotTable.PivotFields("Сумма по полю СрЧек").Caption = "СрЧек (по приемам)";


			pivotTable.AddDataField(pivotTable.PivotFields("Refid"), "Кол-во созданных направлений", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Кол-во приемов").NumberFormat = "# ##0";

			pivotTable.AddDataField(pivotTable.PivotFields("Schedid"), "Кол-во записей (по направлениям)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Кол-во записей (по направлениям)").NumberFormat = "# ##0";

			pivotTable.AddDataField(pivotTable.PivotFields("Treatcode2"), "Кол-во приемов (по направлениям)", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Кол-во приемов (по направлениям)").NumberFormat = "# ##0";

			pivotTable.CalculatedFields().Add("% Исполнено направлений_", "=Treatcode2/Refid", true);
			pivotTable.PivotFields("% Исполнено направлений_").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю % Исполнено направлений_").NumberFormat = "0,00%";
			pivotTable.PivotFields("Сумма по полю % Исполнено направлений_").Caption = "% Исполнено направлений";

			pivotTable.PivotFields("Отделение").ShowDetail = false;
			pivotTable.PivotFields("Филиал").ShowDetail = false;
			//pivotTable.PivotFields("Дата").ShowDetail = false;

			wb.ShowPivotTableFieldList = false;
			//pivotTable.DisplayFieldCaptions = false;

			wsPivote.Range["A1"].Select();
		}
	}
}