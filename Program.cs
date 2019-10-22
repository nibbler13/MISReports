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
using System.Windows;

namespace MISReports {
	public class Program {
		public static string AssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\";

		private static ItemReport itemReport;

		private static DataTable dataTableMainData = null;
		private static DataTable dataTableAverageCheckPreviousWeek = null;
		private static DataTable dataTableAverageCheckPreviousYear = null;
		private static DataTable dataTableWorkLoadA6 = null;
		private static DataTable dataTableWorkloadA11_10 = null;
		private static DataTable dataTableUniqueServiceTotal = null;
		private static DataTable dataTableUniqueServiceLab = null;
		private static DataTable dataTableUniqueServiceLabTotal = null;

		private static DateTime? dateBeginOriginal = null;

		private static Dictionary<string, object> parameters;
		private static Dictionary<string, object> parametersAverageCheckPreviousWeek;
		private static Dictionary<string, object> parametersAverageCheckPreviousYear;

		private static ExcelHandlers.AverageCheck.ItemAverageCheck itemAverageCheckPreviousWeek = null;
		private static ExcelHandlers.AverageCheck.ItemAverageCheck itemAverageCheckPreviousYear = null;
		private static ExcelHandlers.CompetitiveGroups.ItemCompetitiveGroups ItemCompetitiveGroups = null;

		private static string dateBeginStr = string.Empty;
		private static string dateEndStr = string.Empty;
		private static string subject = string.Empty;
		private static string subjectAverageCheckPreviousWeek = string.Empty;
		private static string subjectAverageCheckPreviousYear = string.Empty;
		private static string body = string.Empty;
		private static bool hasError = false;

        private static string fileToUpload = string.Empty;
		private static readonly string mailCopy = Properties.Settings.Default.MailCopy;
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

		private static Tuple<string, string, string>[] licenseStatisticsDBs = 
			new Tuple<string, string, string>[] {
				 new Tuple<string, string, string>("172.16.9.9", "Central", "99_ЦБД"),
				 new Tuple<string, string, string>("172.16.225.2", "mssu", "12_Сущевcкий Вал"),
				 new Tuple<string, string, string>("172.16.225.2", "mskt", "06_Кутузовский"),
				 new Tuple<string, string, string>("172.16.210.203", "web", "Расписание для сайта"),
				 new Tuple<string, string, string>("172.16.225.2", "msn", "02_Мясницкая"),
				 new Tuple<string, string, string>("172.16.9.10", "central-report", "99_ЦБД (Отчеты}"),
				 new Tuple<string, string, string>("172.16.190.6", "dentbase", "01_МДМ"),
				 new Tuple<string, string, string>("172.16.203.2", "spb", "03_СП-Б"),
				 new Tuple<string, string, string>("172.16.127.2", "sretenka", "05_Сретенка"),
				 new Tuple<string, string, string>("172.17.5.2", "snp", "07_Скорая"),
				 new Tuple<string, string, string>("172.16.3.2", "krasn", "08_Краснодар"),
				 new Tuple<string, string, string>("172.16.153.2", "ufa", "09_УФА"),
				 new Tuple<string, string, string>("172.16.158.2", "kazan", "10_Казань"),
				 new Tuple<string, string, string>("172.17.3.2", "yekuk", "15_К-Уральский"),
				 new Tuple<string, string, string>("172.17.100.2", "sctrk", "17_Сочи"),
				 new Tuple<string, string, string>("172.17.10.2", "call_center", "97_Информационный центр")
		};

		public static void Main(string[] args) {
			Logging.ToLog("Старт");

			if (args.Length < 2 || args.Length > 3) {
				Logging.ToLog("Неверное количество параметров");
				WriteOutAcceptedParameters();
				return;
			}

			string reportName = args[0];
			itemReport = new ItemReport(reportName);
			if (!itemReport.IsSettingsLoaded) {
				Logging.ToLog("Неизвестное название отчета: " + reportName);
				WriteOutAcceptedParameters();
				return;
			}

			ParseDateInterval(args);

			if (itemReport.DateBegin == null || itemReport.DateEnd == null) {
				Logging.ToLog("Не удалось распознать временные интервалы формирования отчета");
				WriteOutAcceptedParameters();
				return;
			}

			CreateReport(itemReport);
		}

		public static void CreateReport(ItemReport itemReportToCreate) {
			itemReport = itemReportToCreate;

			FirebirdClient firebirdClient = new FirebirdClient(
				Properties.Settings.Default.MisDbAddress,
				Properties.Settings.Default.MisDbName,
				Properties.Settings.Default.MisDbUser,
				Properties.Settings.Default.MisDbPassword);

			LoadData(firebirdClient);

			firebirdClient.Close();

			WriteDataToFile();

			if (hasError) {
				Logging.ToLog(body);
				itemReport.SetMailTo(mailCopy);
				itemReport.FileResult = string.Empty;
			}

			SaveSettings();

			if (Debugger.IsAttached)
				return;

			if (!string.IsNullOrEmpty(itemReport.FolderToSave))
				SaveReportToFolder();

			if (itemReport.UploadToServer)
				UploadFile();

			if (Logging.bw != null)
				if (MessageBox.Show("Отправить сообщение с отчетом следующим адресатам?" +
					Environment.NewLine + Environment.NewLine + itemReport.MailTo,
					"Отправка сообщения", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
					return;

			string[] attachments;

			if (itemReport.Type == ReportsInfo.Type.AverageCheck)
				attachments = new string[] { itemReport.FileResult, itemReport.FileResultAverageCheckPreviousYear };
			else
				attachments = new string[] { itemReport.FileResult };

			SystemMail.SendMail(subject, body, itemReport.MailTo, attachments);
			Logging.ToLog("Завершение работы");

			return;
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
			DateTime? dateBegin = null;
			DateTime? dateEnd = null;

			if (args.Length == 2) {
				if (args[1].Equals("PreviousMonth")) {
					dateBegin = DateTime.Now.AddMonths(-1).AddDays(-1 * (DateTime.Now.Day - 1));
					dateEnd = dateBegin.Value.AddDays(
						DateTime.DaysInMonth(dateBegin.Value.Year, dateBegin.Value.Month) - 1);
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
			} else
				return;

			if (dateBegin.HasValue && dateEnd.HasValue)
				itemReport.SetPeriod(dateBegin.Value, dateEnd.Value);
		}


		private static void LoadData(FirebirdClient firebirdClient) {
			dateBeginOriginal = itemReport.DateBegin;
			itemReport.SetPeriod(itemReport.DateBegin.AddDays((-1 * itemReport.DateBegin.Day) + 1), itemReport.DateEnd);

			dateBeginStr = dateBeginOriginal.Value.ToShortDateString();
			dateEndStr = itemReport.DateEnd.ToShortDateString();
			subject = ReportsInfo.AcceptedParameters[itemReport.Type] + " с " + dateBeginStr + " по " + dateEndStr;
			Logging.ToLog(subject);

			if (itemReport.Type == ReportsInfo.Type.RegistryMarks)
				dateBeginStr = "01.09.2018";

			if (itemReport.Type == ReportsInfo.Type.MESUsage) {
				Logging.ToLog("Получение данных из базы МИС Инфоклиника за период с " + itemReport.DateBegin.ToShortDateString() + " по " + dateEndStr);
				for (int i = 0; itemReport.DateBegin.AddDays(i) <= itemReport.DateEnd; i++) {
					string dayToGetData = itemReport.DateBegin.AddDays(i).ToShortDateString();
					Logging.ToLog("Получение данных за день: " + dayToGetData);

					Dictionary<string, object> parametersMes = new Dictionary<string, object>() {
						{ "@dateBegin", dayToGetData },
						{ "@dateEnd", dayToGetData }
					};

					DataTable dataTablePart = firebirdClient.GetDataTable(itemReport.SqlQuery, parametersMes);

					if (dataTableMainData == null)
						dataTableMainData = dataTablePart;
					else
						dataTableMainData.Merge(dataTablePart);
				}

				return;
			}

			if (itemReport.Type == ReportsInfo.Type.LicenseStatistics) {
				dataTableMainData = new DataTable();
				dataTableMainData.Columns.Add(new DataColumn("DB", typeof(string)));
				dataTableMainData.Columns.Add(new DataColumn("DATE", typeof(DateTime)));
				dataTableMainData.Columns.Add(new DataColumn("COUNT", typeof(int)));
				foreach (Tuple<string, string, string> item in licenseStatisticsDBs) {
					string dbName = item.Item1 + ":" + item.Item2 + "@" + item.Item3;
					Logging.ToLog("Получение данных из бд: " + dbName);
					try {
						firebirdClient = new FirebirdClient(
							item.Item1, 
							item.Item2, 
							Properties.Settings.Default.MisDbUser, 
							Properties.Settings.Default.MisDbPassword);

						DataTable dataTable = firebirdClient.GetDataTable(itemReport.SqlQuery, new Dictionary<string, object>());
						dataTableMainData.Rows.Add(new object[] {dbName, DateTime.Now, dataTable.Rows[0][0] });
					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
						dataTableMainData.Rows.Add(new object[] { dbName, DateTime.Now, -1 });
					}
				}

				return;
			}

			parameters = new Dictionary<string, object>() {
				{ "@dateBegin", dateBeginStr },
				{ "@dateEnd", dateEndStr }
			};

			Logging.ToLog("Получение данных из базы МИС Инфоклиника за период с " + dateBeginStr + " по " + dateEndStr);

			if (itemReport.Type == ReportsInfo.Type.Workload) {
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

			if (itemReport.Type == ReportsInfo.Type.UniqueServices ||
				itemReport.Type == ReportsInfo.Type.UniqueServicesRegions) {
				string sqlQueryUniqueServiceLab = Properties.Settings.Default.MisDbSqlGetUniqueServicesLab;

				if (itemReport.Type == ReportsInfo.Type.UniqueServicesRegions)
					sqlQueryUniqueServiceLab = Properties.Settings.Default.MisDbSqlGetUniqueServicesRegionsLab;

				dataTableUniqueServiceLab = firebirdClient.GetDataTable(sqlQueryUniqueServiceLab, parameters);

				Dictionary<string, object> parametersTotal = new Dictionary<string, object>() {
					{"@dateBegin",  DateTime.Parse("01.01." + itemReport.DateEnd.ToString("yyyy")).ToShortDateString() },
					{"@dateEnd", dateEndStr }
				};

				dataTableUniqueServiceTotal = firebirdClient.GetDataTable(itemReport.SqlQuery, parametersTotal);
				dataTableUniqueServiceLabTotal = firebirdClient.GetDataTable(sqlQueryUniqueServiceLab, parametersTotal);
			}

			dataTableMainData = firebirdClient.GetDataTable(itemReport.SqlQuery, parameters);
			Logging.ToLog("Получено строк: " + dataTableMainData.Rows.Count);

			if (itemReport.Type == ReportsInfo.Type.PriceListToSite) {
				if (!Directory.Exists(itemReport.FolderToSave)) {
					Logging.ToLog("!!! Не удается получить доступ к папке: " + itemReport.FolderToSave);
					return;
				}

				string priceListToSiteSettingFile = "_Параметры обработки.xlsx";
				string priceListToSiteSettingFilePath = Path.Combine(itemReport.FolderToSave, priceListToSiteSettingFile);
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

            if (itemReport.Type == ReportsInfo.Type.FssInfo)
                ExcelHandlers.FssInfo.PerformData(ref dataTableMainData);

			if (itemReport.Type == ReportsInfo.Type.AverageCheck) {
				int reportWeekNumber = GetIso8601WeekOfYear(dateBeginOriginal.Value);
				int previousWeekNumber = GetIso8601WeekOfYear(dateBeginOriginal.Value.Date.AddDays(-1));

				if (previousWeekNumber > reportWeekNumber)
					subjectAverageCheckPreviousWeek = ReportsInfo.AcceptedParameters[itemReport.Type] + 
						", неделя " + reportWeekNumber + " " + 
						DateTime.Now.Year + " и неделя " + previousWeekNumber + " год " + (DateTime.Now.Year - 1);
				else
					subjectAverageCheckPreviousWeek = ReportsInfo.AcceptedParameters[itemReport.Type] + 
						", недели " + reportWeekNumber + ", " + 
						previousWeekNumber + " год " + DateTime.Now.Year;

				subjectAverageCheckPreviousYear = ReportsInfo.AcceptedParameters[itemReport.Type] + 
					", неделя " + reportWeekNumber + " год " + 
					DateTime.Now.Year + ", " + (DateTime.Now.Year - 1);

				double totalDays = (itemReport.DateEnd - dateBeginOriginal.Value).TotalDays;

				#region previous week
				parametersAverageCheckPreviousWeek = new Dictionary<string, object> {
					{ "@dateBegin", dateBeginOriginal.Value.AddDays(
						-1 * (totalDays + 1)).ToShortDateString() },
					{ "@dateEnd", dateBeginOriginal.Value.AddDays(-1).ToShortDateString() }
				};

				Logging.ToLog("Получение данных из базы МИС Инфоклиника за период с " +
					parametersAverageCheckPreviousWeek["@dateBegin"] +
					" по " + parametersAverageCheckPreviousWeek["@dateEnd"]);
				dataTableAverageCheckPreviousWeek = firebirdClient.GetDataTable(
					itemReport.SqlQuery, parametersAverageCheckPreviousWeek);
				//dataTableAverageCheckPreviousWeek = dataTableMainData.Clone(); 
				Logging.ToLog("Получено строк: " + dataTableAverageCheckPreviousWeek.Rows.Count);

				itemAverageCheckPreviousWeek = ExcelHandlers.AverageCheck.PerformData(dataTableMainData, dataTableAverageCheckPreviousWeek);
				#endregion


				#region previous year
				DateTime previousYearWeekFirstDay = FirstDateOfWeekISO8601(dateBeginOriginal.Value.AddYears(-1).Year, reportWeekNumber);

				parametersAverageCheckPreviousYear = new Dictionary<string, object> {
					{ "@dateBegin", previousYearWeekFirstDay.ToShortDateString()},
					{ "@dateEnd", previousYearWeekFirstDay.AddDays(totalDays).ToShortDateString() }
				};

				Logging.ToLog("Получение данных из базы МИС Инфоклиника за период с " +
					parametersAverageCheckPreviousYear["@dateBegin"] +
					" по " + parametersAverageCheckPreviousYear["@dateEnd"]);
				dataTableAverageCheckPreviousYear = firebirdClient.GetDataTable(
					itemReport.SqlQuery, parametersAverageCheckPreviousYear);
				//dataTableAverageCheckPreviousYear = dataTableMainData.Clone();
				Logging.ToLog("Получено строк: " + dataTableAverageCheckPreviousYear.Rows.Count);

				itemAverageCheckPreviousYear = ExcelHandlers.AverageCheck.PerformData(dataTableMainData, dataTableAverageCheckPreviousYear);
				#endregion
			}

			if (itemReport.Type == ReportsInfo.Type.CompetitiveGroups)
				ItemCompetitiveGroups = ExcelHandlers.CompetitiveGroups.PerformData(dataTableMainData);

			if (itemReport.Type == ReportsInfo.Type.TreatmentsDetails)
				ExcelHandlers.TreatmentsDetails.PerformDataTable(ref dataTableMainData);
		}

		private static int GetIso8601WeekOfYear(DateTime time) {
			// Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll 
			// be the same week# as whatever Thursday, Friday or Saturday are,
			// and we always get those right
			DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
			if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday) {
				time = time.AddDays(3);
			}

			// Return the week of our adjusted day
			return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
		}

		private static DateTime FirstDateOfWeekISO8601(int year, int weekOfYear) {
			DateTime jan1 = new DateTime(year, 1, 1);
			int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

			// Use first Thursday in January to get first week of the year as
			// it will never be in Week 52/53
			DateTime firstThursday = jan1.AddDays(daysOffset);
			var cal = CultureInfo.CurrentCulture.Calendar;
			int firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

			var weekNum = weekOfYear;
			// As we're adding days to a date in Week 1,
			// we need to subtract 1 in order to get the right date for week #1
			if (firstWeek == 1) {
				weekNum -= 1;
			}

			// Using the first Thursday as starting week ensures that we are starting in the right year
			// then we add number of weeks multiplied with days
			var result = firstThursday.AddDays(weekNum * 7);

			// Subtract 3 days from Thursday to get Monday, which is the first weekday in ISO8601
			return result.AddDays(-3);
		}

		private static void WriteDataToFile() {
			if (dataTableMainData.Rows.Count > 0 ||
				itemReport.Type.ToString().StartsWith("VIP_")) {
				Logging.ToLog("Запись данных в файл");

				if (itemReport.Type == ReportsInfo.Type.FreeCellsDay ||
					itemReport.Type == ReportsInfo.Type.FreeCellsWeek) {
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

				if (itemReport.Type == ReportsInfo.Type.MESUsage) {
					Dictionary<string, ItemMESUsageTreatment> treatments =
						ParseMESUsageDataTableToTreatments(dataTableMainData);
					itemReport.FileResult = ExcelHandlers.ExcelGeneral.WriteMesUsageTreatmentsToExcel(treatments,
																  subject,
																  itemReport.TemplateFileName);

				} else if (itemReport.Type == ReportsInfo.Type.TelemedicineOnlyIngosstrakh) {
					itemReport.FileResult = ExcelHandlers.ExcelGeneral.WriteDataTableToExcel(dataTableMainData,
														 subject,
														 itemReport.TemplateFileName,
														 type: itemReport.Type);

				} else if (itemReport.Type == ReportsInfo.Type.Workload) {
					for (int i = 0; i < workloadResultFiles.Count; i++) {
						string key = workloadResultFiles.Keys.ElementAt(i);
						Logging.ToLog("Филиал: " + key);

						workloadResultFiles[key] = ExcelHandlers.ExcelGeneral.WriteDataTableToExcel(dataTableWorkLoadA6,
															 subject + " " + key,
															 itemReport.TemplateFileName,
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

				} else if (itemReport.Type == ReportsInfo.Type.Robocalls) {
					itemReport.FileResult = ExcelHandlers.ExcelGeneral.WriteDataTableToTextFile(dataTableMainData,
															subject,
															itemReport.TemplateFileName);

				} else if (itemReport.Type == ReportsInfo.Type.PriceListToSite) {
					itemReport.FileResult = ExcelHandlers.ExcelGeneral.WriteDataTableToExcel(
						dataTableMainData,
						subject,
						itemReport.TemplateFileName,
						type: itemReport.Type);
					fileToUpload = ExcelHandlers.ExcelGeneral.WriteDataTableToTextFile(
						dataTableMainData,
						"BzPriceListToUpload",
						saveAsJson: true);

				} else if (itemReport.Type == ReportsInfo.Type.TimetableBz) {
					fileToUpload = ExcelHandlers.TimetableBz.PerformData(dataTableMainData);

				} else if (itemReport.Type == ReportsInfo.Type.UniqueServices ||
					itemReport.Type == ReportsInfo.Type.UniqueServicesRegions) {
					itemReport.FileResult = ExcelHandlers.UniqueServices.Process(dataTableMainData,
														 dataTableUniqueServiceTotal,
														 dataTableUniqueServiceLab,
														 dataTableUniqueServiceLabTotal,
														 subject,
														 itemReport.TemplateFileName,
														 dateBeginStr + " - " + dateEndStr,
														 itemReport.Type);

				} else if (itemReport.Type == ReportsInfo.Type.AverageCheck) {
					itemReport.FileResult =
						ExcelHandlers.AverageCheck.WriteAverageCheckToExcel(itemAverageCheckPreviousWeek,
							subjectAverageCheckPreviousWeek, itemReport.TemplateFileName);
					itemReport.FileResultAverageCheckPreviousYear =
						ExcelHandlers.AverageCheck.WriteAverageCheckToExcel(itemAverageCheckPreviousYear,
							subjectAverageCheckPreviousYear, itemReport.TemplateFileName);

				} else if (itemReport.Type == ReportsInfo.Type.CompetitiveGroups) {
					itemReport.FileResult =
						ExcelHandlers.CompetitiveGroups.WriteAverageCheckToExcel(
							ItemCompetitiveGroups, subject, itemReport.TemplateFileName);

				} else {
					itemReport.FileResult = ExcelHandlers.ExcelGeneral.WriteDataTableToExcel(dataTableMainData,
														 subject,
														 itemReport.TemplateFileName,
														 type: itemReport.Type);
				}

				if (File.Exists(itemReport.FileResult) || itemReport.Type == ReportsInfo.Type.Workload) {
					bool isPostProcessingOk = true;

					switch (itemReport.Type) {
						case ReportsInfo.Type.FreeCellsDay:
						case ReportsInfo.Type.FreeCellsWeek:
							isPostProcessingOk = ExcelHandlers.FreeCells.Process(itemReport.FileResult, dateBeginOriginal.Value, itemReport.DateEnd);
							break;

						case ReportsInfo.Type.UnclosedProtocolsWeek:
						case ReportsInfo.Type.UnclosedProtocolsMonth:
							isPostProcessingOk = ExcelHandlers.UnclosedProtocols.Process(itemReport.FileResult);
							break;

						case ReportsInfo.Type.MESUsage:
							isPostProcessingOk = ExcelHandlers.MesUsage.Process(itemReport.FileResult);
							break;

						case ReportsInfo.Type.OnlineAccountsUsage:
							isPostProcessingOk = ExcelHandlers.OnlineAccounts.Process(itemReport.FileResult);
							break;

						case ReportsInfo.Type.TelemedicineOnlyIngosstrakh:
						case ReportsInfo.Type.TelemedicineAll:
							isPostProcessingOk = ExcelHandlers.Telemedicine.Process(itemReport.FileResult);
							break;

						case ReportsInfo.Type.NonAppearance:
							isPostProcessingOk = ExcelHandlers.NonAppearance.Process(itemReport.FileResult, dataTableMainData);
							break;

						case ReportsInfo.Type.VIP_MSSU:
						case ReportsInfo.Type.VIP_Moscow:
						case ReportsInfo.Type.VIP_MSKM:
						case ReportsInfo.Type.VIP_PND:
							isPostProcessingOk = ExcelHandlers.VIP.Process(itemReport.FileResult, itemReport.PreviousFile);
							break;

						case ReportsInfo.Type.RegistryMarks:
							isPostProcessingOk = ExcelHandlers.RegistryMarks.Process(
								itemReport.FileResult, dataTableMainData, dateBeginOriginal.Value);
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
                            isPostProcessingOk = ExcelHandlers.PriceListToSite.Process(itemReport.FileResult);
                            break;

                        case ReportsInfo.Type.GBooking:
						case ReportsInfo.Type.PersonalAccountSchedule:
						case ReportsInfo.Type.ProtocolViewCDBSyncEvent:
							isPostProcessingOk = ExcelHandlers.ExcelGeneral.CopyFormatting(itemReport.FileResult);
							break;

                        case ReportsInfo.Type.FssInfo:
                            isPostProcessingOk = ExcelHandlers.FssInfo.Process(itemReport.FileResult);
                            break;

                        case ReportsInfo.Type.RecordsFromInsuranceCompanies:
                            isPostProcessingOk = ExcelHandlers.RecordsFromInsuranceCompanies.Process(itemReport.FileResult);
                            break;

						case ReportsInfo.Type.AverageCheck:
							isPostProcessingOk = ExcelHandlers.AverageCheck.Process(
								itemReport.FileResult, parameters, parametersAverageCheckPreviousWeek);
							isPostProcessingOk &= ExcelHandlers.AverageCheck.Process(
								itemReport.FileResultAverageCheckPreviousYear, parameters, parametersAverageCheckPreviousYear);
							break;

						case ReportsInfo.Type.CompetitiveGroups:
							isPostProcessingOk = ExcelHandlers.CompetitiveGroups.Process(itemReport.FileResult, parameters);
							break;

						default:
							break;
					}

					if (isPostProcessingOk) {
						body = "Отчет во вложении";
						Logging.ToLog("Данные сохранены в файл: " + (itemReport.Type == ReportsInfo.Type.Workload ?
							string.Join("; ", workloadResultFiles.Values) :
							itemReport.FileResult));
					} else {
						body = "Не удалось выполнить обработку Excel книги";
						hasError = true;
					}
				} else {
					body = "Не удалось записать данные в файл: " + itemReport.FileResult;
					hasError = true;
				}
			} else {
				body = "Отсутствуют данные за период " + itemReport.DateBegin + "-" + itemReport.DateEnd;
				hasError = true;
			}
		}

		private static void SaveReportToFolder() {
			try {
				if (itemReport.Type == ReportsInfo.Type.Workload) {
					Logging.ToLog("Сохранение отчетов в сетевую папку");
					body = "Отчеты сохранены в папку:<br>" + "<a href=\"" + itemReport.FolderToSave + "\">" + itemReport.FolderToSave + "</a><br><br>";
					foreach (KeyValuePair<string, string> pair in workloadResultFiles) {
						Logging.ToLog("Филиал: " + pair.Key);
						if (string.IsNullOrEmpty(pair.Value)) {
							body += pair.Key + ": Нет данных / ошибки обработки<br><br>";
							continue;
						}

						body += pair.Key + ": <br>" +
							SaveFileToNetworkFolder(pair.Value, Path.Combine(itemReport.FolderToSave, pair.Key)) +
							"<br><br>";
					}
				} else {
					body = "Файл с отчетом сохранен по адресу: " + Environment.NewLine +
						SaveFileToNetworkFolder(itemReport.FileResult, itemReport.FolderToSave);
				}
			} catch (Exception e) {
				Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
				body = "Не удалось сохранить отчет в папку " + itemReport.FolderToSave +
					Environment.NewLine + e.Message + Environment.NewLine + e.StackTrace;
				itemReport.SetMailTo(mailCopy);
			}

			itemReport.FileResult = string.Empty;
		}

		public static string SaveFileToNetworkFolder(string localFile, string folderToSave) {
			string fileName = Path.GetFileName(localFile);
			string destFile = Path.Combine(folderToSave, fileName);
			File.Copy(localFile, destFile, true);
			return "<a href=\"" + itemReport.FolderToSave + "\">" + folderToSave + "</a>";
		}

		private static void SaveSettings() {
			switch (itemReport.Type) {
				case ReportsInfo.Type.VIP_MSSU:
					Properties.Settings.Default.PreviousFileVIP_MSSU = itemReport.FileResult;
					break;
				case ReportsInfo.Type.VIP_Moscow:
					Properties.Settings.Default.PreviousFileVIP_Moscow = itemReport.FileResult;
					break;
				case ReportsInfo.Type.VIP_MSKM:
					Properties.Settings.Default.PreviousFileVIP_MSKM = itemReport.FileResult;
					break;
				case ReportsInfo.Type.VIP_PND:
					Properties.Settings.Default.PreviousFileVIP_PND = itemReport.FileResult;
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

			if (itemReport.Type == ReportsInfo.Type.PriceListToSite) {
				url = "https://klinikabudzdorov.ru/export/price/file_input.php";
				method = WebRequestMethods.Http.Post;
			} else if (itemReport.Type == ReportsInfo.Type.TimetableBz) {
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
