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
			OnlineAccountsUsage
		};

		public static Dictionary<ReportType, string> AcceptedParameters = new Dictionary<ReportType, string> {
			{ ReportType.FreeCells, "Отчет по свободным слотам" },
			{ ReportType.UnclosedProtocols, "Отчет по неподписанным протоколам" },
			{ ReportType.MESUsage, "Отчет по использованию МЭС" },
			{ ReportType.OnlineAccountsUsage, "Отчет по записи на прием через личный кабинет" }
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
			} else {
				Logging.ToFile("Неизвестное название отчета: " + reportName);
				WriteOutAcceptedParameters();
				return;
			}

			DateTime? dateBegin = null;
			DateTime? dateEnd = null;

			if (args.Length == 2) {
				if (args[1].Equals("PreviousMonth")) {
					dateBegin = DateTime.Now.AddMonths(-1).AddDays(-1 * (DateTime.Now.Day - 1));
					dateEnd = dateBegin.Value.AddDays(DateTime.DaysInMonth(dateBegin.Value.Year, dateBegin.Value.Month) - 1);
				}
			} else if (args.Length == 3) {
				if (int.TryParse(args[1], out int dateBeginOffset) &&
					int.TryParse(args[2], out int dateEndOffset)) {
					dateBegin = DateTime.Now.AddDays(dateBeginOffset);
					dateEnd = DateTime.Now.AddDays(dateEndOffset);
				} else if (DateTime.TryParseExact(args[1], "dd.MM.yyyy", CultureInfo.InvariantCulture,
					DateTimeStyles.None, out DateTime dateBeginArg) &&
					DateTime.TryParseExact(args[2], "dd.MM.yyyy", CultureInfo.InvariantCulture,
					DateTimeStyles.None, out DateTime dateEndArg)) {
					dateBegin = dateBeginArg;
					dateEnd = dateEndArg;
				}
			}

			if (!dateBegin.HasValue || !dateEnd.HasValue) {
				Logging.ToFile("Не удалось распознать временные интервалы формирования отчета");
				WriteOutAcceptedParameters();
				return;
			}

			FirebirdClient firebirdClient = new FirebirdClient(
				Properties.Settings.Default.MisDbAddress,
				Properties.Settings.Default.MisDbName,
				Properties.Settings.Default.MisDbUser,
				Properties.Settings.Default.MisDbPassword);

			string dateBeginStr = dateBegin.Value.ToShortDateString();
			string dateEndStr = dateEnd.Value.ToShortDateString();
			string subject = AcceptedParameters[reportToCreate] + " с " + dateBeginStr + " по " + dateEndStr;
			Logging.ToFile(subject);

			Dictionary<string, object> parameters = new Dictionary<string, object>() {
				{ "@dateBegin", dateBeginStr },
				{ "@dateEnd", dateEndStr }
			};

			Logging.ToFile("Получение данных из базы МИС Инфоклиника за период с " + dateBeginStr + " по " + dateEndStr);
			DataTable dataTable = firebirdClient.GetDataTable(sqlQuery, parameters);
			Logging.ToFile("Получено строк: " + dataTable.Rows.Count);

			string fileResult = string.Empty;
			string body = string.Empty;
			string mailCopy = Properties.Settings.Default.MailCopy;
			bool hasError = false;

			if (dataTable.Rows.Count > 0) {
				Logging.ToFile("Запись данных в файл Excel");
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
							isPostProcessingOk = NpoiExcelGeneral.PerformFreeCells(fileResult);
							break;
						case ReportType.UnclosedProtocols:
							isPostProcessingOk = NpoiExcelGeneral.PerformUnclosedProtocols(fileResult);
							break;
						case ReportType.MESUsage:
							break;
						case ReportType.OnlineAccountsUsage:
							isPostProcessingOk = NpoiExcelGeneral.PerformOnlineAccountsUsage(fileResult);
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
				body = "Отсутствуют данные за период " + dateBegin + "-" + dateEnd;
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
						treatments[treatcode].ListMES.AddRange(arrayMES);

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
							AGE = row["AGE"].ToString()
						};

						treatment.ListMES.AddRange(arrayMES);

						if (string.IsNullOrEmpty(mid))
							treatment.ListReferralsFromDoc.AddRange(arrayReferrals);
						else
							treatment.ListReferralsFromMes.AddRange(arrayReferrals);

						foreach (string item in arrayAllReferrals) {
							if (!item.Contains(":"))
								continue;

							try {
								string[] referral = item.Split(':');
								string referralName = referral[0];

								if (treatment.ListAllReferrals.ContainsKey(referralName))
									continue;

								int.TryParse(referral[1], out int referralExecuted);
								treatment.ListAllReferrals.Add(referralName, referralExecuted);
							} catch (Exception e) {
								Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
							}
						}

						treatments.Add(treatcode, treatment);
					}
				} catch (Exception e) {
					Logging.ToFile(e.Message);
				}
			}

			return treatments;
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
