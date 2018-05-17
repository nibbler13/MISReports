using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MISReports {
	class Program {
		public static string AssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\";

		public enum ReportType {
			FreeCells,
			UnclosedProtocols,
			MESUsage,
			OnlineAccountsUsage,
			Telemedicine
		};

		public static Dictionary<ReportType, string> AcceptedParameters = new Dictionary<ReportType, string> {
			{ ReportType.FreeCells, "Отчет по свободным слотам" },
			{ ReportType.UnclosedProtocols, "Отчет по неподписанным протоколам" },
			{ ReportType.MESUsage, "Отчет по использованию МЭС" },
			{ ReportType.OnlineAccountsUsage, "Отчет по записи на прием через личный кабинет" },
			{ ReportType.Telemedicine, "Отчет по приемам телемедицины" }
		};

		static void Main(string[] args) {
			Logging.ToFile("Старт");

			if (args.Length < 2 || args.Length > 3) {
				Logging.ToFile("Неверное количество параметров");
				WriteOutAcceptedParameters();
				return;
			}

			string sqlQuery = string.Empty;
			string mailTo = string.Empty;
			string templateFileName = string.Empty;
			ReportType reportToCreate;
			string reportName = args[0];
			if (reportName.Equals(ReportType.FreeCells.ToString())) {
				reportToCreate = ReportType.FreeCells;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetFreeCells;
				mailTo = Properties.Settings.Default.MailToFreeCells;
				templateFileName = Properties.Settings.Default.TemplateFreeCells;
			} else if (reportName.Equals(ReportType.UnclosedProtocols.ToString())) {
				reportToCreate = ReportType.UnclosedProtocols;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetUnclosedProtocols;
				mailTo = Properties.Settings.Default.MailToUnclosedProtocols;
				templateFileName = Properties.Settings.Default.TemplateUnclosedProtocols;
			} else if (reportName.Equals(ReportType.MESUsage.ToString())) {
				reportToCreate = ReportType.MESUsage;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetMESUsage;
				mailTo = Properties.Settings.Default.MailToMESUsage;
				templateFileName = Properties.Settings.Default.TemplateMESUsage;
			} else if (reportName.Equals(ReportType.OnlineAccountsUsage.ToString())) {
				reportToCreate = ReportType.OnlineAccountsUsage;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetOnlineAccountsUsage;
				mailTo = Properties.Settings.Default.MailToOnlineAccountsUsage;
				templateFileName = Properties.Settings.Default.TemplateOnlineAccountsUsage;
			} else if (reportName.Equals(ReportType.Telemedicine.ToString())) {
				reportToCreate = ReportType.Telemedicine;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetTelemedicine;
				templateFileName = Properties.Settings.Default.TemplateTelemedicine;
				mailTo = Properties.Settings.Default.MailToTelemedicine;
			} else {
				Logging.ToFile("Неизвестное название отчета: " + reportName);
				WriteOutAcceptedParameters();
				return;
			}

			DateTime? dateBeginReport = null;
			DateTime? dateEndReport = null;

			if (args.Length == 2) {
				if (args[1].Equals("PreviousMonth")) {
					dateBeginReport = DateTime.Now.AddMonths(-1).AddDays(-1 * (DateTime.Now.Day - 1));
					dateEndReport = dateBeginReport.Value.AddDays(DateTime.DaysInMonth(dateBeginReport.Value.Year, dateBeginReport.Value.Month) - 1);
				}
			} else if (args.Length == 3) {
				if (int.TryParse(args[1], out int dateBeginOffset) &&
					int.TryParse(args[2], out int dateEndOffset)) {
					dateBeginReport = DateTime.Now.AddDays(dateBeginOffset);
					dateEndReport = DateTime.Now.AddDays(dateEndOffset);
				} else if (DateTime.TryParseExact(args[1], "dd.MM.yyyy", CultureInfo.InvariantCulture,
					DateTimeStyles.None, out DateTime dateBeginArg) &&
					DateTime.TryParseExact(args[2], "dd.MM.yyyy", CultureInfo.InvariantCulture,
					DateTimeStyles.None, out DateTime dateEndArg)) {
					dateBeginReport = dateBeginArg;
					dateEndReport = dateEndArg;
				}
			}

			if (!dateBeginReport.HasValue || !dateEndReport.HasValue) {
				Logging.ToFile("Не удалось распознать временные интервалы формирования отчета");
				WriteOutAcceptedParameters();
				return;
			}

			FirebirdClient firebirdClient = new FirebirdClient(
				Properties.Settings.Default.MisDbAddress,
				Properties.Settings.Default.MisDbName,
				Properties.Settings.Default.MisDbUser,
				Properties.Settings.Default.MisDbPassword);

			DateTime? dateBeginOriginal = dateBeginReport;
			dateBeginReport = dateBeginReport.Value.AddDays((-1 * dateBeginReport.Value.Day) + 1);

			string dateBeginStr = dateBeginOriginal.Value.ToShortDateString();
			string dateEndStr = dateEndReport.Value.ToShortDateString();
			string subject = AcceptedParameters[reportToCreate] + " с " + dateBeginStr + " по " + dateEndStr;
			Logging.ToFile(subject);

			Dictionary<string, object> parameters = new Dictionary<string, object>() {
				{ "@dateBegin", dateBeginStr },
				{ "@dateEnd", dateEndStr }
			};

			Logging.ToFile("Получение данных из базы МИС Инфоклиника за период с " + dateBeginReport.Value.ToShortDateString() + " по " + dateEndStr);

			DataTable dataTable = null;
			if (reportToCreate == ReportType.MESUsage ||
				reportToCreate == ReportType.FreeCells) {

				for (int i = 0; dateBeginReport.Value.AddDays(i) <= dateEndReport; i++) {
					string dayToGetData = dateBeginReport.Value.AddDays(i).ToShortDateString();
					Logging.ToFile("Получение данных за день: " + dayToGetData);

					parameters = new Dictionary<string, object>() {
						{ "@dateBegin", dayToGetData },
						{ "@dateEnd", dayToGetData }
					};

					DataTable dataTablePart = firebirdClient.GetDataTable(sqlQuery, parameters);

					if (dataTable == null) {
						dataTable = dataTablePart;
					} else {
						dataTable.Merge(dataTablePart);
					}
				}
			} else {
				dataTable = firebirdClient.GetDataTable(sqlQuery, parameters);
			}

			Logging.ToFile("Получено строк: " + dataTable.Rows.Count);

			string fileResult = string.Empty;
			string body = string.Empty;
			string mailCopy = Properties.Settings.Default.MailCopy;
			bool hasError = false;

			if (dataTable.Rows.Count > 0) {
				Logging.ToFile("Запись данных в файл Excel");
				
				if (reportToCreate == ReportType.FreeCells) {
					DataColumn dataColumn = dataTable.Columns.Add("SortingOrder", typeof(int));
					dataColumn.SetOrdinal(0);

					foreach (DataRow dataRow in dataTable.Rows) {
						int order = 99;

						switch (dataRow["SHORTNAME"].ToString().ToUpper()) {
							case "СУЩ":
								order = 1;
								break;
							case "М-СРЕТ":
								order = 2;
								break;
							case "МДМ":
								order = 3;
								break;
							case "С-ПБ.":
								order = 4;
								break;
							case "УФА":
								order = 5;
								break;
							case "КАЗАНЬ":
								order = 6;
								break;
							case "КРАСН":
								order = 7;
								break;
							case "К-УРАЛ":
								order = 8;
								break;
							case "СОЧИ":
								order = 9;
								break;
							default:
								break;
						}

						dataRow["SortingOrder"] = order;
					}
				}
				
				if (reportToCreate == ReportType.MESUsage) {
					Dictionary<string, ItemMESUsageTreatment> treatments = ParseMESUsageDataTableToTreatments(dataTable);
					fileResult = NpoiExcelGeneral.WriteMesUsageTreatmentsToExcel(treatments, subject, templateFileName);
				} else {
					fileResult = NpoiExcelGeneral.WriteDataTableToExcel(dataTable, subject, templateFileName);
				}

				if (File.Exists(fileResult)) {
					bool isPostProcessingOk = true;

					switch (reportToCreate) {
						case ReportType.FreeCells:
							isPostProcessingOk = NpoiExcelGeneral.PerformFreeCells(fileResult, dateBeginOriginal.Value, dateEndReport.Value);
							break;
						case ReportType.UnclosedProtocols:
							isPostProcessingOk = NpoiExcelGeneral.PerformUnclosedProtocols(fileResult);
							break;
						case ReportType.MESUsage:
							break;
						case ReportType.OnlineAccountsUsage:
							isPostProcessingOk = NpoiExcelGeneral.PerformOnlineAccountsUsage(fileResult);
							break;
						case ReportType.Telemedicine:
							isPostProcessingOk = NpoiExcelGeneral.PerformTelemedicine(fileResult);
							break;
						default:
							break;
					}

					if (isPostProcessingOk) {
						body = "Отчет во вложении";
						Logging.ToFile("Данные сохранены в файл: " + fileResult);
					} else {
						body = "Не удалось выполнить обработку Excel книги";
						hasError = true;
					}
				} else {
					body = "Не удалось записать данные в файл: " + fileResult;
					hasError = true;
				}
			} else {
				body = "Отсутствуют данные за период " + dateBeginReport + "-" + dateEndReport;
				hasError = true;
			}

			if (hasError) {
				Logging.ToFile(body);
				mailTo = mailCopy;
				fileResult = string.Empty;
			}
			
			firebirdClient.Close();

			//if (Debugger.IsAttached)
			//	return;

			SystemMail.SendMail(subject, body, mailTo, fileResult);
			Logging.ToFile("Завершение работы");
		}

		public static Dictionary<string, ItemMESUsageTreatment> ParseMESUsageDataTableToTreatments(DataTable dataTable) {
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
							PAYMENT_TYPE = string.IsNullOrEmpty(row["GRNAME"].ToString()) ? "Страховая компания \\ Безнал" : "Наличный расчет" 
						};
						
						if (string.IsNullOrEmpty(mid))
							treatment.ListReferralsFromDoc.AddRange(arrayReferrals);
						else
							treatment.ListReferralsFromMes.AddRange(arrayReferrals);

						treatment.DictMES = ParseMes(arrayMES);
						treatment.DictAllReferrals = ParseAllReferrals(arrayAllReferrals);
						treatments.Add(treatcode, treatment);
					}
				} catch (Exception e) {
					Logging.ToFile(e.Message);
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
					Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
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
					Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
				}
			}

			return keyValuePairs;
		}

		private static void WriteOutAcceptedParameters() {
			string message = Environment.NewLine + "Формат указания параметров:" + Environment.NewLine +
				"НазваниеОтчета СмещениеДатаНачала СмещениеДатаОкончания (пример: 'FreeCells 0 6')" + Environment.NewLine +
				"НазваниеОтчета ДатаНачала ДатаОкончания (пример: 'FreeCells 01.01.2018 31.01.2018')" +
				"НазваниеОтчета PreviousMonth (пример: 'FreeCells PreviousMonth' - отчет за предыдущий месяц)" +
				Environment.NewLine + Environment.NewLine +
				"Варианты отчетов:" + Environment.NewLine;
			foreach (KeyValuePair<ReportType, string> pair in AcceptedParameters)
				message += pair.Key + " (" + pair.Value + ")" + Environment.NewLine;

			Logging.ToFile(message);
		}
	}
}
