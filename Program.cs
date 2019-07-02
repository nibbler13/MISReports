using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MISReports {
	class Program {
		public static string AssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\";

		private static ReportsInfo.Type reportToCreate;

		private static string sqlQuery = string.Empty;
		private static string folderToSave = string.Empty;
		private static string templateFileName = string.Empty;
		private static string previousFile = string.Empty;
		private static string mailTo = string.Empty;

		private static DateTime? dateBeginReport = null;
		private static DateTime? dateEndReport = null;

		private static DataTable dataTableMainData = null;
		private static DataTable dataTableWorkLoadA6 = null;
		private static DataTable dataTableWorkloadA11_10 = null;
		private static DataTable dataTableUniqueServiceTotal = null;
		private static DataTable dataTableUniqueServiceLab = null;
		private static DataTable dataTableUniqueServiceLabTotal = null;

		private static DateTime? dateBeginOriginal = null;

		private static string dateBeginStr = string.Empty;
		private static string dateEndStr = string.Empty;
		private static string subject = string.Empty;

		private static string fileResult = string.Empty;
        private static string fileToUpload = string.Empty;
		private static readonly string mailCopy = Properties.Settings.Default.MailCopy;
		private static bool hasError = false;
		private static string body = string.Empty;
        private static bool uploadToServer = false;
        private static string priceListToSiteEmptyFields = string.Empty;

		private static readonly Dictionary<string, string> workloadResultFiles = new Dictionary<string, string> {
			{ "_Общий", string.Empty },
			{ "Казань", string.Empty },
			{ "Красн", string.Empty },
			{ "К-УРАЛ", string.Empty },
			{ "МДМ", string.Empty },
			{ "М-СРЕТ", string.Empty },
			{ "Сочи", string.Empty },
			{ "С-Пб", string.Empty },
			{ "СУЩ", string.Empty },
			{ "Уфа", string.Empty }
		};




		static void Main(string[] args) {
			Logging.ToLog("Старт");

			if (args.Length < 2 || args.Length > 3) {
				Logging.ToLog("Неверное количество параметров");
				WriteOutAcceptedParameters();
				return;
			}

			string reportName = args[0];
			if (!LoadSettings(reportName)) {
				Logging.ToLog("Неизвестное название отчета: " + reportName);
				WriteOutAcceptedParameters();
				return;
			}

			ParseDateInterval(args);

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

			LoadData(firebirdClient);

			firebirdClient.Close();

			Logging.ToLog("Получено строк: " + dataTableMainData.Rows.Count);

			WriteDataToFile();

			if (hasError) {
				Logging.ToLog(body);
				mailTo = mailCopy;
				fileResult = string.Empty;
			}

			SaveSettings();

			if (!string.IsNullOrEmpty(folderToSave))
				SaveReportToFolder();

            if (uploadToServer)
                UploadFile();

			if (Debugger.IsAttached)
				return;

			SystemMail.SendMail(subject, body, mailTo, fileResult);
			Logging.ToLog("Завершение работы");
		}




		private static void WriteOutAcceptedParameters() {
			string message = Environment.NewLine + "Формат указания параметров:" + Environment.NewLine +
				"НазваниеОтчета СмещениеДатаНачала СмещениеДатаОкончания (пример: 'FreeCells 0 6')" + Environment.NewLine +
				"НазваниеОтчета ДатаНачала ДатаОкончания (пример: 'FreeCells 01.01.2018 31.01.2018')" +
				"НазваниеОтчета PreviousMonth (пример: 'FreeCells PreviousMonth' - отчет за предыдущий месяц)" +
				Environment.NewLine + Environment.NewLine +
				"Варианты отчетов:" + Environment.NewLine;
			foreach (KeyValuePair<ReportsInfo.Type, string> pair in ReportsInfo.AcceptedParameters)
				message += pair.Key + " (" + pair.Value + ")" + Environment.NewLine;

			Logging.ToLog(message);
		}

		private static void ParseDateInterval(string[] args) {
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
		}

		private static bool LoadSettings(string reportName) {
            if (reportName.Equals(ReportsInfo.Type.FreeCellsDay.ToString())) {
                reportToCreate = ReportsInfo.Type.FreeCellsDay;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetFreeCells;
                mailTo = Properties.Settings.Default.MailToFreeCellsDay;
                templateFileName = Properties.Settings.Default.TemplateFreeCells;

            } else if (reportName.Equals(ReportsInfo.Type.FreeCellsWeek.ToString())) {
                reportToCreate = ReportsInfo.Type.FreeCellsWeek;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetFreeCells;
                mailTo = Properties.Settings.Default.MailToFreeCellsWeek;
                templateFileName = Properties.Settings.Default.TemplateFreeCells;

            } else if (reportName.Equals(ReportsInfo.Type.UnclosedProtocolsWeek.ToString())) {
                reportToCreate = ReportsInfo.Type.UnclosedProtocolsWeek;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetUnclosedProtocols;
                mailTo = Properties.Settings.Default.MailToUnclosedProtocolsWeek;
                templateFileName = Properties.Settings.Default.TemplateUnclosedProtocols;
                folderToSave = Properties.Settings.Default.FolderToSaveUnclosedProtocols;

            } else if (reportName.Equals(ReportsInfo.Type.UnclosedProtocolsMonth.ToString())) {
                reportToCreate = ReportsInfo.Type.UnclosedProtocolsMonth;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetUnclosedProtocols;
                mailTo = Properties.Settings.Default.MailToUnclosedProtocolsMonth;
                templateFileName = Properties.Settings.Default.TemplateUnclosedProtocols;
                folderToSave = Properties.Settings.Default.FolderToSaveUnclosedProtocols;

            } else if (reportName.Equals(ReportsInfo.Type.MESUsage.ToString())) {
                reportToCreate = ReportsInfo.Type.MESUsage;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetMESUsage;
                mailTo = Properties.Settings.Default.MailToMESUsage;
                templateFileName = Properties.Settings.Default.TemplateMESUsage;
                folderToSave = Properties.Settings.Default.FolderToSaveMESUsage;

            } else if (reportName.Equals(ReportsInfo.Type.OnlineAccountsUsage.ToString())) {
                reportToCreate = ReportsInfo.Type.OnlineAccountsUsage;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetOnlineAccountsUsage;
                mailTo = Properties.Settings.Default.MailToOnlineAccountsUsage;
                templateFileName = Properties.Settings.Default.TemplateOnlineAccountsUsage;

            } else if (reportName.Equals(ReportsInfo.Type.TelemedicineOnlyIngosstrakh.ToString())) {
                reportToCreate = ReportsInfo.Type.TelemedicineOnlyIngosstrakh;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetTelemedicine;
                templateFileName = Properties.Settings.Default.TemplateTelemedicine;
                mailTo = Properties.Settings.Default.MailToTelemedicineOnlyIngosstrakh;

            } else if (reportName.Equals(ReportsInfo.Type.TelemedicineAll.ToString())) {
                reportToCreate = ReportsInfo.Type.TelemedicineAll;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetTelemedicine;
                templateFileName = Properties.Settings.Default.TemplateTelemedicine;
                mailTo = Properties.Settings.Default.MailToTelemedicineAll;

            } else if (reportName.Equals(ReportsInfo.Type.NonAppearance.ToString())) {
                reportToCreate = ReportsInfo.Type.NonAppearance;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetNonAppearance;
                templateFileName = Properties.Settings.Default.TemplateNonAppearance;
                mailTo = Properties.Settings.Default.MailToNonAppearance;
                folderToSave = Properties.Settings.Default.FolderToSaveNonAppearance;

            } else if (reportName.Equals(ReportsInfo.Type.VIP_MSSU.ToString())) {
                reportToCreate = ReportsInfo.Type.VIP_MSSU;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetVIP.Replace("@filialList", "12");
                templateFileName = Properties.Settings.Default.TemplateVIP;
                mailTo = Properties.Settings.Default.MailToVIP_MSSU;
                previousFile = Properties.Settings.Default.PreviousFileVIP_MSSU;

            } else if (reportName.Equals(ReportsInfo.Type.VIP_Moscow.ToString())) {
                reportToCreate = ReportsInfo.Type.VIP_Moscow;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetVIP.Replace("@filialList", "1,5,12,6");
                templateFileName = Properties.Settings.Default.TemplateVIP;
                mailTo = Properties.Settings.Default.MailToVIP_Moscow;
                previousFile = Properties.Settings.Default.PreviousFileVIP_Moscow;

            } else if (reportName.Equals(ReportsInfo.Type.VIP_MSKM.ToString())) {
                reportToCreate = ReportsInfo.Type.VIP_MSKM;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetVIP.Replace("@filialList", "1");
                templateFileName = Properties.Settings.Default.TemplateVIP;
                mailTo = Properties.Settings.Default.MailToVIP_MSKM;
                previousFile = Properties.Settings.Default.PreviousFileVIP_MSKM;

            } else if (reportName.Equals(ReportsInfo.Type.VIP_PND.ToString())) {
                reportToCreate = ReportsInfo.Type.VIP_PND;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetVIP.Replace("@filialList", "6");
                templateFileName = Properties.Settings.Default.TemplateVIP;
                mailTo = Properties.Settings.Default.MailToVIP_PND;
                previousFile = Properties.Settings.Default.PreviousFileVIP_PND;

            } else if (reportName.Equals(ReportsInfo.Type.RegistryMarks.ToString())) {
                reportToCreate = ReportsInfo.Type.RegistryMarks;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetRegistryMarks;
                templateFileName = Properties.Settings.Default.TemplateRegistryMarks;
                mailTo = Properties.Settings.Default.MailToRegistryMarks;

            } else if (reportName.Equals(ReportsInfo.Type.Workload.ToString())) {
                reportToCreate = ReportsInfo.Type.Workload;
                templateFileName = Properties.Settings.Default.TemplateWorkload;
                mailTo = Properties.Settings.Default.MailToWorkload;
                folderToSave = Properties.Settings.Default.FolderToSaveWorkload;

            } else if (reportName.Equals(ReportsInfo.Type.Robocalls.ToString())) {
                reportToCreate = ReportsInfo.Type.Robocalls;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetRobocalls;
                templateFileName = Properties.Settings.Default.TemplateRobocalls;
                mailTo = Properties.Settings.Default.MailToRobocalls;

            } else if (reportName.Equals(ReportsInfo.Type.UniqueServices.ToString())) {
                reportToCreate = ReportsInfo.Type.UniqueServices;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetUniqueServices;
                templateFileName = Properties.Settings.Default.TemplateUniqueServices;
                mailTo = Properties.Settings.Default.MailToUniqueServices;

            } else if (reportName.Equals(ReportsInfo.Type.UniqueServicesRegions.ToString())) {
                reportToCreate = ReportsInfo.Type.UniqueServicesRegions;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetUniqueServicesRegions;
                templateFileName = Properties.Settings.Default.TemplateUniqueServicesRegions;
                mailTo = Properties.Settings.Default.MailToUniqueServicesRegions;

            } else if (reportName.Equals(ReportsInfo.Type.PriceListToSite.ToString())) {
                reportToCreate = ReportsInfo.Type.PriceListToSite;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetPriceListToSite;
                templateFileName = Properties.Settings.Default.TemplatePriceListToSite;
                mailTo = Properties.Settings.Default.MailToPriceListToSite;
                folderToSave = Properties.Settings.Default.FolderToSavePriceListToSite;
                uploadToServer = true;

            } else if (reportName.Equals(ReportsInfo.Type.GBooking.ToString())) {
                reportToCreate = ReportsInfo.Type.GBooking;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetGBooking;
                templateFileName = Properties.Settings.Default.TemplateGBooking;
                mailTo = Properties.Settings.Default.MailToGBooking;

            } else if (reportName.Equals(ReportsInfo.Type.PersonalAccountSchedule.ToString())) {
                reportToCreate = ReportsInfo.Type.PersonalAccountSchedule;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetPersonalAccountSchedule;
                templateFileName = Properties.Settings.Default.TemplatePersonalAccountSchedule;
                mailTo = Properties.Settings.Default.MailToPersonalAccountSchedule;

            } else if (reportName.Equals(ReportsInfo.Type.ProtocolViewCDBSyncEvent.ToString())) {
                reportToCreate = ReportsInfo.Type.ProtocolViewCDBSyncEvent;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetProtocolViewCDBSyncEvent;
                templateFileName = Properties.Settings.Default.TemplateProtocolViewCDBSyncEvent;
                mailTo = Properties.Settings.Default.MailToProtocolViewCDBSyncEvent;
                folderToSave = Properties.Settings.Default.FolderToSaveProtocolViewCDBSyncEvent;

            } else if (reportName.Equals(ReportsInfo.Type.FssInfo.ToString())) {
                reportToCreate = ReportsInfo.Type.FssInfo;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetFssInfo;
                templateFileName = Properties.Settings.Default.TemplateFssInfo;
                mailTo = Properties.Settings.Default.MailToFssInfo;

            } else if (reportName.Equals(ReportsInfo.Type.TimetableBz.ToString())) {
                reportToCreate = ReportsInfo.Type.TimetableBz;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetTimetableBz;
                templateFileName = Properties.Settings.Default.TemplateTimetableBz;
                mailTo = Properties.Settings.Default.MailToTimetableBz;
                uploadToServer = true;

            } else if (reportName.Equals(ReportsInfo.Type.RecordsFromInsuranceCompanies.ToString())) {
                reportToCreate = ReportsInfo.Type.RecordsFromInsuranceCompanies;
                sqlQuery = Properties.Settings.Default.MisDbSqlGetRecordsFromInsuranceCompanies;
                templateFileName = Properties.Settings.Default.TemplateRecordsFromInsuranceCompanies;
                mailTo = Properties.Settings.Default.MailToRecordsFromInsuranceCompanies;

            } else
				return false;

			return true;
		}

		private static void LoadData(FirebirdClient firebirdClient) {
			dateBeginOriginal = dateBeginReport;
			dateBeginReport = dateBeginReport.Value.AddDays((-1 * dateBeginReport.Value.Day) + 1);

			dateBeginStr = dateBeginOriginal.Value.ToShortDateString();
			dateEndStr = dateEndReport.Value.ToShortDateString();
			subject = ReportsInfo.AcceptedParameters[reportToCreate] + " с " + dateBeginStr + " по " + dateEndStr;
			Logging.ToLog(subject);

			if (reportToCreate == ReportsInfo.Type.RegistryMarks)
				dateBeginStr = "01.09.2018";

			if (reportToCreate == ReportsInfo.Type.MESUsage) {
				Logging.ToLog("Получение данных из базы МИС Инфоклиника за период с " + dateBeginReport.Value.ToShortDateString() + " по " + dateEndStr);
				for (int i = 0; dateBeginReport.Value.AddDays(i) <= dateEndReport; i++) {
					string dayToGetData = dateBeginReport.Value.AddDays(i).ToShortDateString();
					Logging.ToLog("Получение данных за день: " + dayToGetData);

					Dictionary<string, object> parametersMes = new Dictionary<string, object>() {
						{ "@dateBegin", dayToGetData },
						{ "@dateEnd", dayToGetData }
					};

					DataTable dataTablePart = firebirdClient.GetDataTable(sqlQuery, parametersMes);

					if (dataTableMainData == null)
						dataTableMainData = dataTablePart;
					else
						dataTableMainData.Merge(dataTablePart);
				}

				return;
			}

			Dictionary<string, object> parameters = new Dictionary<string, object>() {
				{ "@dateBegin", dateBeginStr },
				{ "@dateEnd", dateEndStr }
			};

			Logging.ToLog("Получение данных из базы МИС Инфоклиника за период с " + dateBeginStr + " по " + dateEndStr);

			if (reportToCreate == ReportsInfo.Type.Workload) {
				parameters = new Dictionary<string, object>();

				string queryA6 = Path.Combine(AssemblyDirectory, Properties.Settings.Default.QueryWorkloadA6);
				string queryA8_2 = Path.Combine(AssemblyDirectory, Properties.Settings.Default.QueryWorkloadA8_2);
				string queryA11_10 = Path.Combine(AssemblyDirectory, Properties.Settings.Default.QueryWorkloadA11_10);

				if (File.Exists(queryA6) && File.Exists(queryA8_2) && File.Exists(queryA11_10)) {
					try {
						queryA6 = File.ReadAllText(queryA6).Replace("@dateBegin", "'" + dateBeginStr + "'").Replace("@dateEnd", "'" + dateEndStr + "'");
						queryA8_2 = File.ReadAllText(queryA8_2).Replace("@dateBegin", "'" + dateBeginStr + "'").Replace("@dateEnd", "'" + dateEndStr + "'");
						queryA11_10 = File.ReadAllText(queryA11_10).Replace("@dateBegin", "'" + dateBeginStr + "'").Replace("@dateEnd", "'" + dateEndStr + "'");

						dataTableMainData = firebirdClient.GetDataTable(queryA8_2, parameters);
						dataTableWorkLoadA6 = firebirdClient.GetDataTable(queryA6, parameters);
						Logging.ToLog("Получено строк A6: " + dataTableWorkLoadA6.Rows.Count);
						dataTableWorkloadA11_10 = firebirdClient.GetDataTable(queryA11_10, parameters);
						Logging.ToLog("Получено строк A11_10: " + dataTableWorkloadA11_10.Rows.Count);
					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					}
				}

				return;
			}

			if (reportToCreate == ReportsInfo.Type.UniqueServices ||
				reportToCreate == ReportsInfo.Type.UniqueServicesRegions) {
				string sqlQueryUniqueServiceLab = Properties.Settings.Default.MisDbSqlGetUniqueServicesLab;

				if (reportToCreate == ReportsInfo.Type.UniqueServicesRegions)
					sqlQueryUniqueServiceLab = Properties.Settings.Default.MisDbSqlGetUniqueServicesRegionsLab;

				dataTableUniqueServiceLab = firebirdClient.GetDataTable(sqlQueryUniqueServiceLab, parameters);

				Dictionary<string, object> parametersTotal = new Dictionary<string, object>() {
					{"@dateBegin",  DateTime.Parse("01.01." + dateEndReport.Value.ToString("yyyy")).ToShortDateString() },
					{"@dateEnd", dateEndStr }
				};

				dataTableUniqueServiceTotal = firebirdClient.GetDataTable(sqlQuery, parametersTotal);
				dataTableUniqueServiceLabTotal = firebirdClient.GetDataTable(sqlQueryUniqueServiceLab, parametersTotal);
			}

			dataTableMainData = firebirdClient.GetDataTable(sqlQuery, parameters);

			if (reportToCreate == ReportsInfo.Type.PriceListToSite) {
				if (!Directory.Exists(folderToSave)) {
					Logging.ToLog("!!! Не удается получить доступ к папке: " + folderToSave);
					return;
				}

				string priceListToSiteSettingFile = "_Параметры обработки.xlsx";
				string priceListToSiteSettingFilePath = Path.Combine(folderToSave, priceListToSiteSettingFile);
				if (!File.Exists(priceListToSiteSettingFilePath)) {
					Logging.ToLog("!!! Не удается получить доступ к файлу с настройками: " + priceListToSiteSettingFilePath);
					return;
				}

				string sheetNameExclusions = "Исключения";
				string sheetNameGrouping = "Группировки";
				string sheetNamePriorities = "Приоритеты";

				Logging.ToLog("Считывание настроек из файла: " + priceListToSiteSettingFilePath);

				try {
					DataTable dataTablePriceExclusions = ExcelHandlers.ExcelGeneral.ReadExcelFile(priceListToSiteSettingFilePath, sheetNameExclusions);
					Logging.ToLog("Считано строк: " + dataTablePriceExclusions.Rows.Count);
					DataTable dataTablePriceGrouping = ExcelHandlers.ExcelGeneral.ReadExcelFile(priceListToSiteSettingFilePath, sheetNameGrouping);
					Logging.ToLog("Считано строк: " + dataTablePriceGrouping.Rows.Count);
					DataTable dataTablePricePriorities = ExcelHandlers.ExcelGeneral.ReadExcelFile(priceListToSiteSettingFilePath, sheetNamePriorities);
					Logging.ToLog("Считано строк: " + dataTablePricePriorities.Rows.Count);

					dataTableMainData = ExcelHandlers.PriceListToSite.PerformData(
						dataTableMainData, dataTablePriceExclusions, dataTablePriceGrouping, dataTablePricePriorities, out priceListToSiteEmptyFields);
				} catch (Exception e) {
					Logging.ToLog(e.StackTrace + Environment.NewLine + e.StackTrace);
					return;
				}
			}

            if (reportToCreate == ReportsInfo.Type.FssInfo)
                ExcelHandlers.FssInfo.PerformData(ref dataTableMainData);
		}

		private static void WriteDataToFile() {
			if (dataTableMainData.Rows.Count > 0 ||
				reportToCreate.ToString().StartsWith("VIP_")) {
				Logging.ToLog("Запись данных в файл");

				if (reportToCreate == ReportsInfo.Type.FreeCellsDay ||
					reportToCreate == ReportsInfo.Type.FreeCellsWeek) {
					DataColumn dataColumn = dataTableMainData.Columns.Add("SortingOrder", typeof(int));
					dataColumn.SetOrdinal(0);

					foreach (DataRow dataRow in dataTableMainData.Rows) {
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

                if (reportToCreate == ReportsInfo.Type.MESUsage) {
                    Dictionary<string, ItemMESUsageTreatment> treatments =
                        ParseMESUsageDataTableToTreatments(dataTableMainData);
                    fileResult = ExcelHandlers.ExcelGeneral.WriteMesUsageTreatmentsToExcel(treatments,
                                                                  subject,
                                                                  templateFileName);

                } else if (reportToCreate == ReportsInfo.Type.TelemedicineOnlyIngosstrakh) {
                    fileResult = ExcelHandlers.ExcelGeneral.WriteDataTableToExcel(dataTableMainData,
                                                         subject,
                                                         templateFileName,
                                                         type: reportToCreate);

                } else if (reportToCreate == ReportsInfo.Type.Workload) {
                    for (int i = 0; i < workloadResultFiles.Count; i++) {
                        string key = workloadResultFiles.Keys.ElementAt(i);
                        Logging.ToLog("Филиал: " + key);

                        workloadResultFiles[key] = ExcelHandlers.ExcelGeneral.WriteDataTableToExcel(dataTableWorkLoadA6,
                                                             subject + " " + key,
                                                             templateFileName,
                                                             "Услуги Мет. 1",
                                                             true,
                                                             key);

                        if (string.IsNullOrEmpty(workloadResultFiles[key]))
                            continue;

                        ExcelHandlers.ExcelGeneral.WriteDataTableToExcel(dataTableWorkloadA11_10,
                                                subject,
                                                workloadResultFiles[key],
                                                "Искл. услуги",
                                                false,
                                                key);

                        ExcelHandlers.ExcelGeneral.WriteDataTableToExcel(dataTableMainData,
                                                subject,
                                                workloadResultFiles[key],
                                                "Расчет",
                                                false,
                                                key);
                    }

                } else if (reportToCreate == ReportsInfo.Type.Robocalls) {
                    fileResult = ExcelHandlers.ExcelGeneral.WriteDataTableToTextFile(dataTableMainData,
                                                            subject,
                                                            templateFileName);

                } else if (reportToCreate == ReportsInfo.Type.PriceListToSite) {
                    fileResult = ExcelHandlers.ExcelGeneral.WriteDataTableToExcel(
                        dataTableMainData,
                        subject,
                        templateFileName,
                        type: reportToCreate);
                    fileToUpload = ExcelHandlers.ExcelGeneral.WriteDataTableToTextFile(
                        dataTableMainData, 
                        "BzPriceListToUpload", 
                        saveAsJson: true);

                } else if (reportToCreate == ReportsInfo.Type.TimetableBz) {
                    fileToUpload = ExcelHandlers.TimetableBz.PerformData(dataTableMainData);

                } else if (reportToCreate == ReportsInfo.Type.UniqueServices ||
					reportToCreate == ReportsInfo.Type.UniqueServicesRegions) {
					fileResult = ExcelHandlers.UniqueServices.Process(dataTableMainData,
														 dataTableUniqueServiceTotal,
														 dataTableUniqueServiceLab,
														 dataTableUniqueServiceLabTotal,
														 subject,
														 templateFileName,
														 dateBeginStr + " - " + dateEndStr,
														 reportToCreate);

				} else {
					fileResult = ExcelHandlers.ExcelGeneral.WriteDataTableToExcel(dataTableMainData,
														 subject,
														 templateFileName,
														 type: reportToCreate);
				}

				if (File.Exists(fileResult) || reportToCreate == ReportsInfo.Type.Workload) {
					bool isPostProcessingOk = true;

					switch (reportToCreate) {
						case ReportsInfo.Type.FreeCellsDay:
						case ReportsInfo.Type.FreeCellsWeek:
							isPostProcessingOk = ExcelHandlers.FreeCells.Process(fileResult, dateBeginOriginal.Value, dateEndReport.Value);
							break;
						case ReportsInfo.Type.UnclosedProtocolsWeek:
						case ReportsInfo.Type.UnclosedProtocolsMonth:
							isPostProcessingOk = ExcelHandlers.UnclosedProtocols.Process(fileResult);
							break;
						case ReportsInfo.Type.MESUsage:
							isPostProcessingOk = ExcelHandlers.MesUsage.Process(fileResult);
							break;
						case ReportsInfo.Type.OnlineAccountsUsage:
							isPostProcessingOk = ExcelHandlers.OnlineAccounts.Process(fileResult);
							break;
						case ReportsInfo.Type.TelemedicineOnlyIngosstrakh:
						case ReportsInfo.Type.TelemedicineAll:
							isPostProcessingOk = ExcelHandlers.Telemedicine.Process(fileResult);
							break;
						case ReportsInfo.Type.NonAppearance:
							isPostProcessingOk = ExcelHandlers.NonAppearance.Process(fileResult, dataTableMainData);
							break;
						case ReportsInfo.Type.VIP_MSSU:
						case ReportsInfo.Type.VIP_Moscow:
						case ReportsInfo.Type.VIP_MSKM:
						case ReportsInfo.Type.VIP_PND:
							isPostProcessingOk = ExcelHandlers.VIP.Process(fileResult, previousFile);
							break;
						case ReportsInfo.Type.RegistryMarks:
							isPostProcessingOk = ExcelHandlers.RegistryMarks.Process(
								fileResult, dataTableMainData, dateBeginOriginal.Value);
							break;
						case ReportsInfo.Type.Workload:
							bool isAllOk = true;
							Logging.ToLog("Пост-обработка");
							foreach (string currentFileResult in workloadResultFiles.Values) {
								Logging.ToLog("Файл: " + currentFileResult);

								if (string.IsNullOrEmpty(currentFileResult))
									continue;

								if (!ExcelHandlers.Workload.Process(currentFileResult))
									isAllOk = false;
							}

							isPostProcessingOk = isAllOk;
							break;
                        case ReportsInfo.Type.PriceListToSite:
                            isPostProcessingOk = ExcelHandlers.PriceListToSite.Process(fileResult);
                            break;
                        case ReportsInfo.Type.GBooking:
						case ReportsInfo.Type.PersonalAccountSchedule:
						case ReportsInfo.Type.ProtocolViewCDBSyncEvent:
							isPostProcessingOk = ExcelHandlers.ExcelGeneral.CopyFormatting(fileResult);
							break;
                        case ReportsInfo.Type.FssInfo:
                            isPostProcessingOk = ExcelHandlers.FssInfo.Process(fileResult);
                            break;
                        case ReportsInfo.Type.RecordsFromInsuranceCompanies:
                            isPostProcessingOk = ExcelHandlers.RecordsFromInsuranceCompanies.Process(fileResult);
                            break;
						default:
							break;
					}

					if (isPostProcessingOk) {
						body = "Отчет во вложении";
						Logging.ToLog("Данные сохранены в файл: " + (reportToCreate == ReportsInfo.Type.Workload ?
							string.Join("; ", workloadResultFiles.Values) :
							fileResult));
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
		}

		private static void SaveReportToFolder() {
			try {
				if (reportToCreate == ReportsInfo.Type.Workload) {
					Logging.ToLog("Сохранение отчетов в сетевую папку");
					body = "Отчеты сохранены в папку:<br>" + "<a href=\"" + folderToSave + "\">" + folderToSave + "</a><br><br>";
					foreach (KeyValuePair<string, string> pair in workloadResultFiles) {
						Logging.ToLog("Филиал: " + pair.Key);
						if (string.IsNullOrEmpty(pair.Value)) {
							body += pair.Key + ": Нет данных / ошибки обработки<br><br>";
							continue;
						}

						body += pair.Key + ": <br>" +
							SaveFileToNetworkFolder(pair.Value, Path.Combine(folderToSave, pair.Key)) +
							"<br><br>";
					}
				} else {
					body = "Файл с отчетом сохранен по адресу: " + Environment.NewLine +
						SaveFileToNetworkFolder(fileResult, folderToSave);
				}
			} catch (Exception e) {
				Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
				body = "Не удалось сохранить отчет в папку " + folderToSave +
					Environment.NewLine + e.Message + Environment.NewLine + e.StackTrace;
				mailTo = mailCopy;
			}

			fileResult = string.Empty;
		}

		public static string SaveFileToNetworkFolder(string localFile, string folderToSave) {
			string fileName = Path.GetFileName(localFile);
			string destFile = Path.Combine(folderToSave, fileName);
			File.Copy(localFile, destFile, true);
			return "<a href=\"" + folderToSave + "\">" + folderToSave + "</a>";
		}

		private static void SaveSettings() {
			switch (reportToCreate) {
				case ReportsInfo.Type.VIP_MSSU:
					Properties.Settings.Default.PreviousFileVIP_MSSU = fileResult;
					break;
				case ReportsInfo.Type.VIP_Moscow:
					Properties.Settings.Default.PreviousFileVIP_Moscow = fileResult;
					break;
				case ReportsInfo.Type.VIP_MSKM:
					Properties.Settings.Default.PreviousFileVIP_MSKM = fileResult;
					break;
				case ReportsInfo.Type.VIP_PND:
					Properties.Settings.Default.PreviousFileVIP_PND = fileResult;
					break;
				default:
					break;
			}

			Properties.Settings.Default.Save();
		}

        private static bool PostDataToServer() {
			string aFileurl = fileToUpload;
			string aTargetUrl = "ftp://prodoctorov.ru" + "/" + "bz_timetable.json";
			Debug.WriteLine("creating ftp upload. Source: " + aFileurl + " Target: " + aTargetUrl);
			System.IO.FileStream aFileStream = null;
			System.IO.Stream aRequestStream = null;

			try {
				var aFtpClient = (FtpWebRequest)FtpWebRequest.Create(aTargetUrl);
				aFtpClient.Credentials = new NetworkCredential("bud-zdorov-moskva-3846", "ef5febfa506709e7788e925122dc1106");
				aFtpClient.Method = WebRequestMethods.Ftp.UploadFile;
				aFtpClient.UseBinary = true;
				aFtpClient.KeepAlive = true;
				aFtpClient.UsePassive = true;
				aFtpClient.Proxy = null;

				var aFileInfo = new System.IO.FileInfo(aFileurl);
				aFtpClient.ContentLength = aFileInfo.Length;
				byte[] aBuffer = new byte[4097];
				int aBytes = 0;
				int aTotal_bytes = (int)aFileInfo.Length;
				aFileStream = aFileInfo.OpenRead();
				aRequestStream = aFtpClient.GetRequestStream();
				while (aTotal_bytes > 0) {
					aBytes = aFileStream.Read(aBuffer, 0, aBuffer.Length);
					aRequestStream.Write(aBuffer, 0, aBytes);
					aTotal_bytes = aTotal_bytes - aBytes;
				}
				aFileStream.Close();
				aRequestStream.Close();
				var uploadResponse = (FtpWebResponse)aFtpClient.GetResponse();
				Debug.WriteLine(uploadResponse.StatusDescription);
				uploadResponse.Close();
				return true;
			} catch (Exception e) {
				if (aFileStream != null) aFileStream.Close();
				if (aRequestStream != null) aRequestStream.Close();

				Debug.WriteLine(e.ToString());
				return false;
			}
		}

        private static void UploadFile() {
            string msg = "Загрузка файла на сервер";
            Logging.ToLog(msg);
            body += Environment.NewLine + Environment.NewLine + msg;

            string url = string.Empty;
            string user = string.Empty;
            string password = string.Empty;
            string method = string.Empty;

			if (reportToCreate == ReportsInfo.Type.PriceListToSite) {
				url = "https://klinikabudzdorov.ru/export/price/file_input.php";
				method = WebRequestMethods.Http.Post;
			} else if (reportToCreate == ReportsInfo.Type.TimetableBz) {
				PostDataToServer();
				return;
			} else {
				Logging.ToLog("Не заданы параметры, возврат");
				return;
			}

            try {
                using (WebClient client = new WebClient()) {
                    if (!string.IsNullOrEmpty(user) && 
                        !string.IsNullOrEmpty(password))
                        client.Credentials = new NetworkCredential(user, password);

                    byte[] responseArray = client.UploadFile(url, method, fileToUpload);
                    string response = System.Text.Encoding.GetEncoding("windows-1252").GetString(responseArray);
                    Logging.ToLog(response);

                    body += Environment.NewLine + response;

                    if (!string.IsNullOrEmpty(priceListToSiteEmptyFields))
                        body += Environment.NewLine + Environment.NewLine +
                            "Услуги с недостающими данными: " + Environment.NewLine +
                            priceListToSiteEmptyFields;
                }
            } catch (Exception e) {
                hasError = true;
                msg = e.Message + Environment.NewLine + e.StackTrace;
                Logging.ToLog(msg);
                body += Environment.NewLine + msg;
            }

            try {
                File.Delete(fileToUpload);
            } catch (Exception e) {
                Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
            }
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

	}
}
