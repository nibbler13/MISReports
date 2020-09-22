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
				string histnum = row["HISTNUM"].ToString();
				string phone = row["PHONE"].ToString();
				string depname = row["DEPNAME"].ToString();
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
					Histnum = histnum,
					Phone = phone,
					TreatType = services.ToLower().Contains("первичный") ? "Первичный" : (services.ToLower().Contains("повторный") ? "Повторный" : ""),
					Depname = depname,
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
			public string Histnum { get; set; }
			public string Phone { get; set; }
			public string TreatType { get; set; }
			public string Depname { get; set; }
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
					treatment.Histnum,
					treatment.Phone,
					treatment.TreatType,
					treatment.Depname,
					treatment.Doc
				};

				int serviceCountResult = 0;
				int serviceQuantityResult = 0;
				double serviceAmountResult = 0;

				int serviceRow = rowNumber;
				foreach (ItemService service in treatment.Services) {
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
				foreach (ItemReferral referral in treatment.Referrals.Values) {
					referralQuantityResult++;

					foreach (ItemService service in referral.Services) {
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
							referral.Refid,
							service.Name,
							service.Kodoper,
							service.Count,
							service.Amount
						};

						if (referral.Schedule is null)
							valuesService.AddRange(new List<object> {
								"Отсутствует",
								string.Empty,
								string.Empty,
								string.Empty,
								string.Empty,
								string.Empty,
								string.Empty
							});
						else {
							referralScheduleQuantityResult++;

							valuesService.AddRange(new List<object> {
								referral.Schedule.Workdate,
								referral.Schedule.Schedid,
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
								string.Empty,
								string.Empty,
								string.Empty,
								string.Empty,
								string.Empty,
								string.Empty
							});
						else {
							referralTreatQuantityResult++;

							valuesService.AddRange(new List<object> {
								"Да",
								referral.Treat.Treatdate,
								referral.Treat.Treatcode
							});

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

								treatService.IsUsed = true;
							}
						}

						valuesToWrite.AddRange(valuesService);
						WriteOutRow(valuesToWrite, sheet, referralRow, columnNumber);
						referralRow++;
					}
				}

				columnNumber = 0;
				rowNumber = serviceRow > referralRow ? serviceRow : referralRow;

				List<object> valuesToWriteResult = new List<object>();
				valuesToWriteResult.AddRange(valuesTreatment);
				valuesToWriteResult.Add("Итого по услугам");
				valuesToWriteResult.Add(serviceQuantityResult + " шт.");
				valuesToWriteResult.Add(serviceCountResult);
				valuesToWriteResult.Add(serviceAmountResult);

				if (referralQuantityResult > 0) {
					valuesToWriteResult.Add("Итого по направлениям");
					valuesToWriteResult.Add(referralQuantityResult + " шт.");
				} else {
					valuesToWriteResult.Add("Отсутствуют");
					valuesToWriteResult.Add(string.Empty);
				}

				if (referralServiceQuantityResult > 0) {
					valuesToWriteResult.Add("Итого по услугам");
					valuesToWriteResult.Add(referralServiceQuantityResult + " шт.");
					valuesToWriteResult.Add(referralServiceCountResult);
					valuesToWriteResult.Add(referralServiceAmountResult);
				} else {
					valuesToWriteResult.AddRange(new List<object>{
						string.Empty,
						string.Empty,
						string.Empty,
						string.Empty 
					});
				}

				if (referralScheduleQuantityResult > 0) {
					valuesToWriteResult.Add("Итого по назначениям в расписание");
					valuesToWriteResult.Add(referralScheduleQuantityResult + " шт.");
					valuesToWriteResult.AddRange(new List<object> {
						string.Empty,
						string.Empty,
						string.Empty,
						string.Empty,
						string.Empty
					});
				} else {
					if (referralQuantityResult > 0)
						valuesToWriteResult.Add("Нет назначений");
					else
						valuesToWriteResult.Add(string.Empty);

					valuesToWriteResult.AddRange(new List<object>{
						string.Empty,
						string.Empty,
						string.Empty,
						string.Empty,
						string.Empty,
						string.Empty
					});
				}

				if (referralQuantityResult > 0) {
					string serviceResult = "Нет услуг в направлениях";
					if (referralServiceCountResult > 0)
						serviceResult = ((double)referralTreatServiceCountResult / (double)referralServiceCountResult * 100.0)
							.ToString("N2", CultureInfo.CurrentCulture) + "% - исполнения (по услугам)";
					valuesToWriteResult.Add(serviceResult);

					if (referralTreatQuantityResult > 0) {
						valuesToWriteResult.Add("Итого по приемам");
						valuesToWriteResult.Add(referralTreatQuantityResult + " шт.");
					} else
						valuesToWriteResult.AddRange(new List<object> {
							string.Empty,
							string.Empty
						});

					if (referralTreatServiceQuantityResult > 0) {
						valuesToWriteResult.Add("Итого по услугам");
						valuesToWriteResult.Add(referralTreatServiceQuantityResult + " шт.");
						valuesToWriteResult.Add(referralTreatServiceCountResult);
						valuesToWriteResult.Add(referralTreatServiceAmountResult);
					} else
						valuesToWriteResult.AddRange(new List<object> {
							string.Empty,
							string.Empty,
							string.Empty,
							string.Empty
						});
				} else {
					if (referralQuantityResult > 0)
						valuesToWriteResult.Add("Нет приемов");
					else
						valuesToWriteResult.Add(string.Empty);

					valuesToWriteResult.AddRange(new List<object>{
						string.Empty,
						string.Empty,
						string.Empty,
						string.Empty,
						string.Empty,
						string.Empty
					});
				}

				if (referralServiceAmountResult > 0)
					valuesToWriteResult.Add(
						((double)referralTreatServiceAmountResult / (double)referralServiceAmountResult * 100.0d)
						.ToString("N2", CultureInfo.CurrentCulture));
				else
					valuesToWriteResult.Add("Нет направлений");

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
					else if (DateTime.TryParseExact(value.ToString(), "dd.MM.yyyy h:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
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
				bool isDark = false;
				for (int i = 2; i < usedRows; i++) {
					object tr = ws.Range["C" + i].Value2;
					string treatcode = tr is null ? string.Empty : tr.ToString();

					for (int y = i + 1; y <= usedRows; y++) {
						object trNext = ws.Range["C" + y].Value2;
						string treatcodeNext = trNext is null ? string.Empty : trNext.ToString();
						if (treatcodeNext.Equals(treatcode) && y != usedRows)
							continue;

						SetColorForRange(ws.Range["A" + i + ":AH" + (y - 1)], isDark);
						isDark = !isDark;
						i = y - 1;
						break;
					}
				}
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static void SetColorForRange(Excel.Range rangeToColorize, bool isDark) {
			double tintAndShade = -1;
			if (isDark)
				tintAndShade = 0.599993896298105;
			else 
				tintAndShade = 0.799981688894314;

			if (tintAndShade != -1) {
				rangeToColorize.Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
				rangeToColorize.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4;
				rangeToColorize.Interior.TintAndShade = tintAndShade;
				rangeToColorize.Interior.PatternTintAndShade = 0;
			}
		}
	}
}