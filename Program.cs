﻿using System;
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
			FreeCellsDay,
			FreeCellsWeek,
			UnclosedProtocolsWeek,
			UnclosedProtocolsMonth,
			MESUsage,
			OnlineAccountsUsage,
			TelemedicineOnlyIngosstrakh,
			TelemedicineAll,
			NonAppearance,
			VIP_MSSU,
			VIP_Moscow,
			VIP_MSKM,
			VIP_PND,
			RegistryMarks,
			Workload,
			Robocalls,
			UniqueServices
		};

		public static Dictionary<ReportType, string> AcceptedParameters = new Dictionary<ReportType, string> {
			{ ReportType.FreeCellsDay, "Отчет по свободным слотам" },
			{ ReportType.FreeCellsWeek, "Отчет по свободным слотам" },
			{ ReportType.UnclosedProtocolsWeek, "Отчет по неподписанным протоколам" },
			{ ReportType.UnclosedProtocolsMonth, "Отчет по неподписанным протоколам" },
			{ ReportType.MESUsage, "Отчет по использованию МЭС" },
			{ ReportType.OnlineAccountsUsage, "Отчет по записи на прием через личный кабинет" },
			{ ReportType.TelemedicineOnlyIngosstrakh, "Отчет по приемам телемедицины - только Ингосстрах" },
			{ ReportType.TelemedicineAll, "Отчет по приемам телемедицины - все типы оплаты" },
			{ ReportType.NonAppearance, "Отчет по неявкам" },
			{ ReportType.VIP_MSSU, "Отчет по ВИП-пациентам Сущевка" },
			{ ReportType.VIP_Moscow, "Отчет по ВИП-пациентам Москва" },
			{ ReportType.VIP_MSKM, "Отчет по ВИП-пациентам Фрунзенская" },
			{ ReportType.VIP_PND, "Отчет по ВИП-пациентам ПНД" },
			{ ReportType.RegistryMarks, "Отчет по оценкам регистратуры" },
			{ ReportType.Workload, "Отчет по загрузке сотрудников" },
			{ ReportType.Robocalls, "Информация для автообзвона" },
			{ReportType.UniqueServices, "Отчет по уникальным услугам" }
		};

		static void Main(string[] args) {
			Logging.ToLog("Старт");

			if (args.Length < 2 || args.Length > 3) {
				Logging.ToLog("Неверное количество параметров");
				WriteOutAcceptedParameters();
				return;
			}

			string sqlQuery = string.Empty;
			string folderToSave = string.Empty;
			string previousFile = string.Empty;
			ReportType reportToCreate;
			string reportName = args[0];
			string templateFileName;
			string mailTo;

			if (reportName.Equals(ReportType.FreeCellsDay.ToString())) {
				reportToCreate = ReportType.FreeCellsDay;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetFreeCells;
				mailTo = Properties.Settings.Default.MailToFreeCellsDay;
				templateFileName = Properties.Settings.Default.TemplateFreeCells;

			} else if (reportName.Equals(ReportType.FreeCellsWeek.ToString())) {
				reportToCreate = ReportType.FreeCellsWeek;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetFreeCells;
				mailTo = Properties.Settings.Default.MailToFreeCellsWeek;
				templateFileName = Properties.Settings.Default.TemplateFreeCells;

			} else if (reportName.Equals(ReportType.UnclosedProtocolsWeek.ToString())) {
				reportToCreate = ReportType.UnclosedProtocolsWeek;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetUnclosedProtocols;
				mailTo = Properties.Settings.Default.MailToUnclosedProtocolsWeek;
				templateFileName = Properties.Settings.Default.TemplateUnclosedProtocols;
				folderToSave = Properties.Settings.Default.FolderToSaveUnclosedProtocols;

			} else if (reportName.Equals(ReportType.UnclosedProtocolsMonth.ToString())) {
				reportToCreate = ReportType.UnclosedProtocolsMonth;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetUnclosedProtocols;
				mailTo = Properties.Settings.Default.MailToUnclosedProtocolsMonth;
				templateFileName = Properties.Settings.Default.TemplateUnclosedProtocols;
				folderToSave = Properties.Settings.Default.FolderToSaveUnclosedProtocols;

			} else if (reportName.Equals(ReportType.MESUsage.ToString())) {
				reportToCreate = ReportType.MESUsage;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetMESUsage;
				mailTo = Properties.Settings.Default.MailToMESUsage;
				templateFileName = Properties.Settings.Default.TemplateMESUsage;
				folderToSave = Properties.Settings.Default.FolderToSaveMESUsage;

			} else if (reportName.Equals(ReportType.OnlineAccountsUsage.ToString())) {
				reportToCreate = ReportType.OnlineAccountsUsage;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetOnlineAccountsUsage;
				mailTo = Properties.Settings.Default.MailToOnlineAccountsUsage;
				templateFileName = Properties.Settings.Default.TemplateOnlineAccountsUsage;

			} else if (reportName.Equals(ReportType.TelemedicineOnlyIngosstrakh.ToString())) {
				reportToCreate = ReportType.TelemedicineOnlyIngosstrakh;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetTelemedicine;
				templateFileName = Properties.Settings.Default.TemplateTelemedicine;
				mailTo = Properties.Settings.Default.MailToTelemedicineOnlyIngosstrakh;

			} else if (reportName.Equals(ReportType.TelemedicineAll.ToString())) {
				reportToCreate = ReportType.TelemedicineAll;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetTelemedicine;
				templateFileName = Properties.Settings.Default.TemplateTelemedicine;
				mailTo = Properties.Settings.Default.MailToTelemedicineAll;

			} else if (reportName.Equals(ReportType.NonAppearance.ToString())) {
				reportToCreate = ReportType.NonAppearance;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetNonAppearance;
				templateFileName = Properties.Settings.Default.TemplateNonAppearance;
				mailTo = Properties.Settings.Default.MailToNonAppearance;
				folderToSave = Properties.Settings.Default.FolderToSaveNonAppearance;

			} else if (reportName.Equals(ReportType.VIP_MSSU.ToString())) {
				reportToCreate = ReportType.VIP_MSSU;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetVIP.Replace("@filialList", "12");
				templateFileName = Properties.Settings.Default.TemplateVIP;
				mailTo = Properties.Settings.Default.MailToVIP_MSSU;
				previousFile = Properties.Settings.Default.PreviousFileVIP_MSSU;

			} else if (reportName.Equals(ReportType.VIP_Moscow.ToString())) {
				reportToCreate = ReportType.VIP_Moscow;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetVIP.Replace("@filialList", "1,5,12,6");
				templateFileName = Properties.Settings.Default.TemplateVIP;
				mailTo = Properties.Settings.Default.MailToVIP_Moscow;
				previousFile = Properties.Settings.Default.PreviousFileVIP_Moscow;

			} else if (reportName.Equals(ReportType.VIP_MSKM.ToString())) {
				reportToCreate = ReportType.VIP_MSKM;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetVIP.Replace("@filialList", "1");
				templateFileName = Properties.Settings.Default.TemplateVIP;
				mailTo = Properties.Settings.Default.MailToVIP_MSKM;
				previousFile = Properties.Settings.Default.PreviousFileVIP_MSKM;

			} else if (reportName.Equals(ReportType.VIP_PND.ToString())) {
				reportToCreate = ReportType.VIP_PND;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetVIP.Replace("@filialList", "6");
				templateFileName = Properties.Settings.Default.TemplateVIP;
				mailTo = Properties.Settings.Default.MailToVIP_PND;
				previousFile = Properties.Settings.Default.PreviousFileVIP_PND;

			} else if (reportName.Equals(ReportType.RegistryMarks.ToString())) {
				reportToCreate = ReportType.RegistryMarks;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetRegistryMarks;
				templateFileName = Properties.Settings.Default.TemplateRegistryMarks;
				mailTo = Properties.Settings.Default.MailToRegistryMarks;

			} else if (reportName.Equals(ReportType.Workload.ToString())) {
				reportToCreate = ReportType.Workload;
				templateFileName = Properties.Settings.Default.TemplateWorkload;
				mailTo = Properties.Settings.Default.MailToWorkload;

			} else if (reportName.Equals(ReportType.Robocalls.ToString())) {
				reportToCreate = ReportType.Robocalls;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetRobocalls;
				templateFileName = Properties.Settings.Default.TemplateRobocalls;
				mailTo = Properties.Settings.Default.MailToRobocalls;

			} else if (reportName.Equals(ReportType.UniqueServices.ToString())) {
				reportToCreate = ReportType.UniqueServices;
				sqlQuery = Properties.Settings.Default.MisDbSqlGetUniqueServices;
				templateFileName = Properties.Settings.Default.TemplateUniqueServices;
				mailTo = Properties.Settings.Default.MailToUniqueServices;

			} else {
				Logging.ToLog("Неизвестное название отчета: " + reportName);
				WriteOutAcceptedParameters();
				return;
			}

			DateTime? dateBeginReport = null;
			DateTime? dateEndReport = null;

			if (args.Length == 2) {
				if (args[1].Equals("PreviousMonth")) {
					dateBeginReport = DateTime.Now.AddMonths(-1).AddDays(-1 * (DateTime.Now.Day - 1));
					dateEndReport = dateBeginReport.Value.AddDays(
						DateTime.DaysInMonth(dateBeginReport.Value.Year, dateBeginReport.Value.Month) - 1);
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
				Logging.ToLog("Не удалось распознать временные интервалы формирования отчета");
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
            Logging.ToLog(subject);

            if (reportToCreate == ReportType.RegistryMarks)
                dateBeginStr = "01.09.2018";

			DataTable dataTable = null;
			DataTable dataTableWorkLoadA6 = null;
			DataTable dataTableWorkloadA11_10 = null;
			DataTable dataTableUniqueServiceTotal = null;

			if (reportToCreate == ReportType.MESUsage) {
				Logging.ToLog("Получение данных из базы МИС Инфоклиника за период с " + dateBeginReport.Value.ToShortDateString() + " по " + dateEndStr);
				for (int i = 0; dateBeginReport.Value.AddDays(i) <= dateEndReport; i++) {
					string dayToGetData = dateBeginReport.Value.AddDays(i).ToShortDateString();
					Logging.ToLog("Получение данных за день: " + dayToGetData);

					Dictionary<string, object> parameters = new Dictionary<string, object>() {
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
				Dictionary<string, object> parameters = new Dictionary<string, object>() {
					{ "@dateBegin", dateBeginStr },
					{ "@dateEnd", dateEndStr }
				};

				Logging.ToLog("Получение данных из базы МИС Инфоклиника за период с " + dateBeginStr + " по " + dateEndStr);

				if (reportToCreate == ReportType.Workload) {
					parameters = new Dictionary<string, object>();

					string queryA6 = Path.Combine(AssemblyDirectory, Properties.Settings.Default.QueryWorkloadA6);
					string queryA8_2 = Path.Combine(AssemblyDirectory, Properties.Settings.Default.QueryWorkloadA8_2);
					string queryA11_10 = Path.Combine(AssemblyDirectory, Properties.Settings.Default.QueryWorkloadA11_10);

					if (File.Exists(queryA6) && File.Exists(queryA8_2) && File.Exists(queryA11_10)) {
						try {
							queryA6 = File.ReadAllText(queryA6).Replace("@dateBegin", "'" + dateBeginStr + "'").Replace("@dateEnd", "'" + dateEndStr + "'");
							queryA8_2 = File.ReadAllText(queryA8_2).Replace("@dateBegin", "'" + dateBeginStr + "'").Replace("@dateEnd", "'" + dateEndStr + "'");
							queryA11_10 = File.ReadAllText(queryA11_10).Replace("@dateBegin", "'" + dateBeginStr + "'").Replace("@dateEnd", "'" + dateEndStr + "'");

							dataTable = firebirdClient.GetDataTable(queryA8_2, parameters);
							dataTableWorkLoadA6 = firebirdClient.GetDataTable(queryA6, parameters);
							Logging.ToLog("Получено строк A6: " + dataTableWorkLoadA6.Rows.Count);
							dataTableWorkloadA11_10 = firebirdClient.GetDataTable(queryA11_10, parameters);
							Logging.ToLog("Получено строк A11_10: " + dataTableWorkloadA11_10.Rows.Count);
						} catch (Exception e) {
							Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
						}
					}
				} else
					dataTable = firebirdClient.GetDataTable(sqlQuery, parameters);

				if (reportToCreate == ReportType.UniqueServices) {
					Dictionary<string, object> parametersTotal = new Dictionary<string, object>() {
						{"@dateBegin",  DateTime.Parse("01.01." + dateEndReport.Value.ToString("yyyy")).ToShortDateString() },
						{"@dateEnd", dateEndStr }
					};

					dataTableUniqueServiceTotal = firebirdClient.GetDataTable(sqlQuery, parametersTotal);
				}

			}

			Logging.ToLog("Получено строк: " + dataTable.Rows.Count);

			string fileResult = string.Empty;
			string mailCopy = Properties.Settings.Default.MailCopy;
			bool hasError = false;
			string body;

			if (dataTable.Rows.Count > 0 ||
				reportToCreate.ToString().StartsWith("VIP_")) {
				Logging.ToLog("Запись данных в файл Excel");

				if (reportToCreate == ReportType.FreeCellsDay ||
					reportToCreate == ReportType.FreeCellsWeek) {
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
					Dictionary<string, ItemMESUsageTreatment> treatments = 
						ParseMESUsageDataTableToTreatments(dataTable);
					fileResult = NpoiExcelGeneral.WriteMesUsageTreatmentsToExcel(treatments,
																  subject,
																  templateFileName);

				} else if (reportToCreate == ReportType.TelemedicineOnlyIngosstrakh) {
					fileResult = NpoiExcelGeneral.WriteDataTableToExcel(dataTable,
														 subject,
														 templateFileName,
														 true);

				} else if (reportToCreate == ReportType.Workload) {
					fileResult = NpoiExcelGeneral.WriteDataTableToExcel(dataTableWorkLoadA6,
														 subject,
														 templateFileName,
														 false,
														 "Услуги Мет. 1");

					NpoiExcelGeneral.WriteDataTableToExcel(dataTableWorkloadA11_10,
											subject,
											fileResult,
											false,
											"Искл. услуги",
											false);

					NpoiExcelGeneral.WriteDataTableToExcel(dataTable,
											subject,
											fileResult,
											false,
											"Расчет",
											false);

				} else if (reportToCreate == ReportType.Robocalls) {
					fileResult = NpoiExcelGeneral.WriteDataTableToTextFile(dataTable,
															subject,
															templateFileName);

				} else if (reportToCreate == ReportType.UniqueServices) {
					fileResult = NpoiExcelGeneral.PerformUniqueServices(dataTable,
														 dataTableUniqueServiceTotal,
														 subject,
														 templateFileName,
														 dateBeginStr + " - " + dateEndStr);

				} else {
					fileResult = NpoiExcelGeneral.WriteDataTableToExcel(dataTable,
														 subject,
														 templateFileName);
				}

				if (File.Exists(fileResult)) {
					bool isPostProcessingOk = true;

					switch (reportToCreate) {
						case ReportType.FreeCellsDay:
						case ReportType.FreeCellsWeek:
							isPostProcessingOk = NpoiExcelGeneral.PerformFreeCells(fileResult, dateBeginOriginal.Value, dateEndReport.Value);
							break;
						case ReportType.UnclosedProtocolsWeek:
						case ReportType.UnclosedProtocolsMonth:
							isPostProcessingOk = NpoiExcelGeneral.PerformUnclosedProtocols(fileResult);
							break;
						case ReportType.MESUsage:
							isPostProcessingOk = NpoiExcelGeneral.PerformMesUsage(fileResult);
							break;
						case ReportType.OnlineAccountsUsage:
							isPostProcessingOk = NpoiExcelGeneral.PerformOnlineAccountsUsage(fileResult);
							break;
						case ReportType.TelemedicineOnlyIngosstrakh:
						case ReportType.TelemedicineAll:
							isPostProcessingOk = NpoiExcelGeneral.PerformTelemedicine(fileResult);
							break;
						case ReportType.NonAppearance:
							isPostProcessingOk = NpoiExcelGeneral.PerformNonAppearance(fileResult, dataTable);
							break;
						case ReportType.VIP_MSSU:
						case ReportType.VIP_Moscow:
						case ReportType.VIP_MSKM:
						case ReportType.VIP_PND:
							isPostProcessingOk = NpoiExcelGeneral.PerformVIP(fileResult, previousFile);
							break;
						case ReportType.RegistryMarks:
							isPostProcessingOk = NpoiExcelGeneral.PerformRegistryMarks(
								fileResult, dataTable, dateBeginOriginal.Value);
							break;
						case ReportType.Workload:
							isPostProcessingOk = NpoiExcelGeneral.PerformWorkload(fileResult);
							break;
						default:
							break;
					}

					if (isPostProcessingOk) {
						body = "Отчет во вложении";
						Logging.ToLog("Данные сохранены в файл: " + fileResult);
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
				Logging.ToLog(body);
				mailTo = mailCopy;
				fileResult = string.Empty;
			}
			
			firebirdClient.Close();

			switch (reportToCreate) {
				case ReportType.VIP_MSSU:
					Properties.Settings.Default.PreviousFileVIP_MSSU = fileResult;
					break;
				case ReportType.VIP_Moscow:
					Properties.Settings.Default.PreviousFileVIP_Moscow = fileResult;
					break;
				case ReportType.VIP_MSKM:
					Properties.Settings.Default.PreviousFileVIP_MSKM = fileResult;
					break;
				case ReportType.VIP_PND:
					Properties.Settings.Default.PreviousFileVIP_PND = fileResult;
					break;
				default:
					break;
			}

			if (!string.IsNullOrEmpty(folderToSave)) {
				try {
					string fileName = Path.GetFileName(fileResult);
					string destFile = Path.Combine(folderToSave, fileName);
					File.Copy(fileResult, destFile, true);
					body = "Файл с отчетом сохранен по адресу: " +
						Environment.NewLine + "<a href=\"" + destFile + "\">" + destFile + "</a>";
				} catch (Exception e) {
					Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
					body = "Не удалось сохранить отчет в папку " + folderToSave +
						Environment.NewLine + e.Message + Environment.NewLine + e.StackTrace;
					mailTo = mailCopy;
				}

				fileResult = string.Empty;
			}

			Properties.Settings.Default.Save();

			if (Debugger.IsAttached)
				return;

			SystemMail.SendMail(subject, body, mailTo, fileResult);
			Logging.ToLog("Завершение работы");
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
					Logging.ToLog(e.Message);
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

		private static void WriteOutAcceptedParameters() {
			string message = Environment.NewLine + "Формат указания параметров:" + Environment.NewLine +
				"НазваниеОтчета СмещениеДатаНачала СмещениеДатаОкончания (пример: 'FreeCells 0 6')" + Environment.NewLine +
				"НазваниеОтчета ДатаНачала ДатаОкончания (пример: 'FreeCells 01.01.2018 31.01.2018')" +
				"НазваниеОтчета PreviousMonth (пример: 'FreeCells PreviousMonth' - отчет за предыдущий месяц)" +
				Environment.NewLine + Environment.NewLine +
				"Варианты отчетов:" + Environment.NewLine;
			foreach (KeyValuePair<ReportType, string> pair in AcceptedParameters)
				message += pair.Key + " (" + pair.Value + ")" + Environment.NewLine;

			Logging.ToLog(message);
		}
	}
}
