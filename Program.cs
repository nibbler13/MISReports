using MISReports.ExcelHandlers;
using NPOI.SS.Formula.Functions;
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

		private static AverageCheck.ItemAverageCheck itemAverageCheckPreviousWeek = null;
		private static AverageCheck.ItemAverageCheck itemAverageCheckPreviousYear = null;
		private static AverageCheck.ItemAverageCheck itemAverageCheckIGS = null;
		private static CompetitiveGroups.ItemCompetitiveGroups ItemCompetitiveGroups = null;

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

		private static Dictionary<string, string> workloadResultFiles = new Dictionary<string, string> {
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

		private static readonly List<ReportsInfo.Type> TreatmentsDetailsType = new List<ReportsInfo.Type> {
			ReportsInfo.Type.TreatmentsDetailsAbsolut,
			ReportsInfo.Type.TreatmentsDetailsAlfa,
			ReportsInfo.Type.TreatmentsDetailsAlfaSpb,
			ReportsInfo.Type.TreatmentsDetailsAlliance,
			ReportsInfo.Type.TreatmentsDetailsBestdoctor,
			ReportsInfo.Type.TreatmentsDetailsEnergogarant,
			ReportsInfo.Type.TreatmentsDetailsIngosstrakhAdult,
			ReportsInfo.Type.TreatmentsDetailsIngosstrakhKid,
			ReportsInfo.Type.TreatmentsDetailsLiberty,
			ReportsInfo.Type.TreatmentsDetailsMetlife,
			ReportsInfo.Type.TreatmentsDetailsRenessans,
			ReportsInfo.Type.TreatmentsDetailsReso,
			ReportsInfo.Type.TreatmentsDetailsRosgosstrakh,
			ReportsInfo.Type.TreatmentsDetailsSmp,
			ReportsInfo.Type.TreatmentsDetailsSogaz,
			ReportsInfo.Type.TreatmentsDetailsSoglasie,
			ReportsInfo.Type.TreatmentsDetailsVsk,
			ReportsInfo.Type.TreatmentsDetailsVtb,
			ReportsInfo.Type.TreatmentsDetailsIngosstrakhSochi,
			ReportsInfo.Type.TreatmentsDetailsIngosstrakhKrasnodar,
			ReportsInfo.Type.TreatmentsDetailsIngosstrakhUfa,
			ReportsInfo.Type.TreatmentsDetailsIngosstrakhSpb,
			ReportsInfo.Type.TreatmentsDetailsIngosstrakhKazan,
			ReportsInfo.Type.TreatmentsDetailsBestDoctorSpb,
			ReportsInfo.Type.TreatmentsDetailsBestDoctorUfa,
			ReportsInfo.Type.TreatmentsDetailsSogazUfa,
		};

		private static readonly Tuple<string, string, string>[] infoclinicaDBs = 
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

		private static readonly Dictionary<string, string> registryMotivationQueries = new Dictionary<string, string> {
			{ "ЛК", "select " + Environment.NewLine +
					"doc.dcode " + Environment.NewLine +
					", doc.fullname " + Environment.NewLine +
					", doc.doctpost " + Environment.NewLine +
					", count(distinct cd.docid )as kol " + Environment.NewLine +
					"--, d.docname as doc_f " + Environment.NewLine +
					", f.shortname  as f_n " + Environment.NewLine +
					"from cldocuments cd " + Environment.NewLine +
					"join DOCTEMPLATES d on  cd.DOCTYPE = d.ID " + Environment.NewLine +
					"join doctor doc on cd.UID = doc.DCODE " + Environment.NewLine +
					"join filials f on doc.FILIAL =  f.FILID " + Environment.NewLine +
					"where " + Environment.NewLine +
					"--cd.DOCTYPE in (990000244,990000254,990000257) " + Environment.NewLine +
					"d.docname containing 'Соглашение на использование' " + Environment.NewLine +
					"and cd.DOCDATE between @dateBegin and @dateEnd " + Environment.NewLine +
					"group by 1,2,3,5 " + Environment.NewLine +
					"order by 2 asc " },
			{ "RG", "select " + Environment.NewLine +
					"doc.dcode " + Environment.NewLine +
					", doc.fullname " + Environment.NewLine +
					", doc.doctpost " + Environment.NewLine +
					", count(distinct cd.docid )as kol " + Environment.NewLine +
					"--, d.docname as doc_f " + Environment.NewLine +
					", f.shortname  as f_n " + Environment.NewLine +
					"from cldocuments cd " + Environment.NewLine +
					"join DOCTEMPLATES d on  cd.DOCTYPE = d.ID " + Environment.NewLine +
					"join doctor doc on cd.UID = doc.DCODE " + Environment.NewLine +
					"join filials f on doc.FILIAL =  f.FILID " + Environment.NewLine +
					"where " + Environment.NewLine +
					"--cd.DOCTYPE in (990000244,990000254,990000257) " + Environment.NewLine +
					"d.docname containing 'RG ' " + Environment.NewLine +
					"and cd.DOCDATE between @dateBegin and @dateEnd " + Environment.NewLine +
					"group by 1,2,3,5 " + Environment.NewLine +
					"order by 2 asc " },
			{ "Выписка", "select " + Environment.NewLine +
					"doc.dcode " + Environment.NewLine +
					", doc.fullname " + Environment.NewLine +
					", doc.doctpost " + Environment.NewLine +
					", count(distinct cd.docid )as kol " + Environment.NewLine +
					"--, d.docname as doc_f " + Environment.NewLine +
					", f.shortname  as f_n " + Environment.NewLine +
					"from cldocuments cd " + Environment.NewLine +
					"join DOCTEMPLATES d on  cd.DOCTYPE = d.ID " + Environment.NewLine +
					"join doctor doc on cd.UID = doc.DCODE " + Environment.NewLine +
					"join filials f on doc.FILIAL =  f.FILID " + Environment.NewLine +
					"where " + Environment.NewLine +
					"--cd.DOCTYPE in (990000244,990000254,990000257) " + Environment.NewLine +
					"d.docname containing 'Выписка по обследованию' " + Environment.NewLine +
					"and cd.DOCDATE between @dateBegin and @dateEnd " + Environment.NewLine +
					"group by 1,2,3,5 " + Environment.NewLine +
					"order by 2 asc " },
			{ "ИДС", "select " + Environment.NewLine +
					"doc.dcode " + Environment.NewLine +
					", doc.fullname " + Environment.NewLine +
					", doc.doctpost " + Environment.NewLine +
					", count(distinct cd.docid )as kol " + Environment.NewLine +
					"--, d.docname as doc_f " + Environment.NewLine +
					", f.shortname  as f_n " + Environment.NewLine +
					"from cldocuments cd " + Environment.NewLine +
					"join DOCTEMPLATES d on  cd.DOCTYPE = d.ID " + Environment.NewLine +
					"join doctor doc on cd.UID = doc.DCODE " + Environment.NewLine +
					"join filials f on doc.FILIAL =  f.FILID " + Environment.NewLine +
					"where " + Environment.NewLine +
					"--cd.DOCTYPE in (990000244,990000254,990000257) " + Environment.NewLine +
					"d.docname containing 'ИДС' " + Environment.NewLine +
					"and cd.DOCDATE between @dateBegin and @dateEnd " + Environment.NewLine +
					"group by 1,2,3,5 " + Environment.NewLine +
					"order by 2 asc " },
			{ "Договор", "select " + Environment.NewLine +
					"doc.dcode " + Environment.NewLine +
					", doc.fullname " + Environment.NewLine +
					", doc.doctpost " + Environment.NewLine +
					", count(distinct cd.docid )as kol " + Environment.NewLine +
					"--, d.docname as doc_f " + Environment.NewLine +
					", f.shortname  as f_n " + Environment.NewLine +
					"from cldocuments cd " + Environment.NewLine +
					"join DOCTEMPLATES d on  cd.DOCTYPE = d.ID " + Environment.NewLine +
					"join doctor doc on cd.UID = doc.DCODE " + Environment.NewLine +
					"join filials f on doc.FILIAL =  f.FILID " + Environment.NewLine +
					"where " + Environment.NewLine +
					"--cd.DOCTYPE in (990000244,990000254,990000257) " + Environment.NewLine +
					"d.docname containing 'Договор' " + Environment.NewLine +
					"and cd.DOCDATE between @dateBegin and @dateEnd " + Environment.NewLine +
					"group by 1,2,3,5 " + Environment.NewLine +
					"order by 2 asc " },
			{ "Заявление карта", "select " + Environment.NewLine +
					"doc.dcode " + Environment.NewLine +
					", doc.fullname " + Environment.NewLine +
					", doc.doctpost " + Environment.NewLine +
					", count(distinct cd.docid )as kol " + Environment.NewLine +
					"--, d.docname as doc_f " + Environment.NewLine +
					", f.shortname  as f_n " + Environment.NewLine +
					"from cldocuments cd " + Environment.NewLine +
					"join DOCTEMPLATES d on  cd.DOCTYPE = d.ID " + Environment.NewLine +
					"join doctor doc on cd.UID = doc.DCODE " + Environment.NewLine +
					"join filials f on doc.FILIAL =  f.FILID " + Environment.NewLine +
					"where " + Environment.NewLine +
					"--cd.DOCTYPE in (990000244,990000254,990000257) " + Environment.NewLine +
					"d.docname containing 'Заявление на выдачу' " + Environment.NewLine +
					"and cd.DOCDATE between @dateBegin and @dateEnd " + Environment.NewLine +
					"group by 1,2,3,5 " + Environment.NewLine +
					"order by 2 asc " },
			{ "Заявление анализы", "select " + Environment.NewLine +
					"doc.dcode " + Environment.NewLine +
					", doc.fullname " + Environment.NewLine +
					", doc.doctpost " + Environment.NewLine +
					", count(distinct cd.docid )as kol " + Environment.NewLine +
					"--, d.docname as doc_f " + Environment.NewLine +
					", f.shortname  as f_n " + Environment.NewLine +
					"from cldocuments cd " + Environment.NewLine +
					"join DOCTEMPLATES d on  cd.DOCTYPE = d.ID " + Environment.NewLine +
					"join doctor doc on cd.UID = doc.DCODE " + Environment.NewLine +
					"join filials f on doc.FILIAL =  f.FILID " + Environment.NewLine +
					"where " + Environment.NewLine +
					"--cd.DOCTYPE in (990000244,990000254,990000257) " + Environment.NewLine +
					"d.docname containing 'Заявление на получение результатов анализов на email' " + Environment.NewLine +
					"and cd.DOCDATE between @dateBegin and @dateEnd " + Environment.NewLine +
					"group by 1,2,3,5 " + Environment.NewLine +
					"order by 2 asc " },
			{ "Заявление возврат", "select " + Environment.NewLine +
					"doc.dcode " + Environment.NewLine +
					", doc.fullname " + Environment.NewLine +
					", doc.doctpost " + Environment.NewLine +
					", count(distinct cd.docid )as kol " + Environment.NewLine +
					"--, d.docname as doc_f " + Environment.NewLine +
					", f.shortname  as f_n " + Environment.NewLine +
					"from cldocuments cd " + Environment.NewLine +
					"join DOCTEMPLATES d on  cd.DOCTYPE = d.ID " + Environment.NewLine +
					"join doctor doc on cd.UID = doc.DCODE " + Environment.NewLine +
					"join filials f on doc.FILIAL =  f.FILID " + Environment.NewLine +
					"where " + Environment.NewLine +
					"--cd.DOCTYPE in (990000244,990000254,990000257) " + Environment.NewLine +
					"d.docname containing 'Заявление на возврат денежных средств' " + Environment.NewLine +
					"and cd.DOCDATE between @dateBegin and @dateEnd " + Environment.NewLine +
					"group by 1,2,3,5 " + Environment.NewLine +
					"order by 2 asc " },
			{ "Отметок", "select " + Environment.NewLine +
					"d.dcode " + Environment.NewLine +
					", d.fullname " + Environment.NewLine +
					", d.doctpost " + Environment.NewLine +
					", count(distinct s.SCHEDID) " + Environment.NewLine +
					", f.shortname " + Environment.NewLine +
					"from Schedule s " + Environment.NewLine +
					"join treat t on t.treatcode = s.treatcode " + Environment.NewLine +
					"join filials f on f.filid = s.filial " + Environment.NewLine +
					"join doctor d on s.VISIT_UID = d.dcode " + Environment.NewLine +
					"where WorkDate between @dateBegin and @dateEnd " + Environment.NewLine +
					"and s.CLVISIT = 1 " + Environment.NewLine +
					"group by 1,2,3,5 " },
			{ "Записи", "select " + Environment.NewLine +
					"d.dcode " + Environment.NewLine +
					", d.fullname " + Environment.NewLine +
					", d.doctpost " + Environment.NewLine +
					", count(distinct s.SCHEDID) " + Environment.NewLine +
					", f.shortname " + Environment.NewLine +
					"from Schedule s " + Environment.NewLine +
					"left join treat t on t.treatcode = s.treatcode " + Environment.NewLine +
					"join filials f on f.filid = s.filial " + Environment.NewLine +
					"join doctor d on s.CREATORID = d.dcode " + Environment.NewLine +
					"where WorkDate between @dateBegin and @dateEnd " + Environment.NewLine +
					"group by 1,2,3,5 " },
			{ "Анализы", "select " + Environment.NewLine +
					"d.dcode " + Environment.NewLine +
					", d.fullname " + Environment.NewLine +
					", d.doctpost " + Environment.NewLine +
					", count(o.lgid) " + Environment.NewLine +
					", f.shortname " + Environment.NewLine +
					"from operlog o " + Environment.NewLine +
					"left join operlogref op on o.eventtype = op.eventtype " + Environment.NewLine +
					"join doctor d on o.uid = d.dcode " + Environment.NewLine +
					"join filials f on d.FILIAL = f.filid " + Environment.NewLine +
					"where o.eventtype in (229) " + Environment.NewLine +
					"and o.eventdate between @dateBegin and @dateEnd " + Environment.NewLine +
					"group by 1,2,3,5 " }
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







		public static void CreateReport(ItemReport itemReportToCreate, bool? needToAskToSend = null) {
			if (Debugger.IsAttached)
				workloadResultFiles = new Dictionary<string, string> { { "_Общий", string.Empty } };

			if (itemReportToCreate.Type == ReportsInfo.Type.TreatmentsDetailsAll) {
				foreach (ReportsInfo.Type type in TreatmentsDetailsType) {
					ItemReport report = new ItemReport(type.ToString());
					report.SetPeriod(itemReportToCreate.DateBegin, itemReportToCreate.DateEnd);
					CreateReport(report, needToAskToSend);

					if (File.Exists(report.FileResult))
						itemReportToCreate.FileResult = report.FileResult;
				}

				return;
			}

			itemReport = itemReportToCreate;

			Logging.ToLog(
				"Создание подключения к БД: " + 
				Properties.Settings.Default.MisDbAddress + ":" + 
				Properties.Settings.Default.MisDbName);

			IDbClient dbClient = new FirebirdClient(
				Properties.Settings.Default.MisDbAddress,
				Properties.Settings.Default.MisDbName,
				Properties.Settings.Default.MisDbUser,
				Properties.Settings.Default.MisDbPassword);

			if (itemReport.UseVerticaDb)
				dbClient = new VerticaClient(
					Properties.Settings.Default.VerticaDbAddress,
					Properties.Settings.Default.VerticaDbDatabase,
					Properties.Settings.Default.VerticaDbUser,
					Properties.Settings.Default.VerticaDbPassword);

			LoadData(dbClient);
			dbClient.Close();
			WriteDataToFile();

			if (itemReport.Type == ReportsInfo.Type.LicenseEndingDates)
				return;

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

			if (needToAskToSend.HasValue && needToAskToSend.Value)
				if (MessageBox.Show("Отправить сообщение с отчетом следующим адресатам?" +
					Environment.NewLine + Environment.NewLine + itemReport.MailTo,
					"Отправка сообщения", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
					return;

			string[] attachments;

			if (itemReport.Type == ReportsInfo.Type.AverageCheckRegular)
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

			DateTime tst = new DateTime(2020, 5, 18);
			Console.WriteLine(tst.AddDays(-17).ToShortDateString());
			Console.WriteLine(tst.AddDays(-3).ToShortDateString());

			if (args.Length == 2) {
				if (args[1].Equals("PreviousMonth")) {
					dateBegin = DateTime.Now.AddMonths(-1).AddDays(-1 * (DateTime.Now.Day - 1));
					dateEnd = dateBegin.Value.AddDays(
						DateTime.DaysInMonth(dateBegin.Value.Year, dateBegin.Value.Month) - 1);
				} else if (args[1].Equals("PreviousMonthSecondPart")) {
					dateBegin = DateTime.Now.AddMonths(-1).AddDays(-1 * (DateTime.Now.Day - 1)).AddDays(15);
					dateEnd = DateTime.Now.AddDays(-1 * DateTime.Now.Day);
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


		private static void LoadData(IDbClient dbClient) {
			dateBeginOriginal = itemReport.DateBegin;
			dateBeginStr = dateBeginOriginal.Value.ToShortDateString();
			dateEndStr = itemReport.DateEnd.ToShortDateString();
			subject = ReportsInfo.AcceptedParameters[itemReport.Type] + " с " + dateBeginStr + " по " + dateEndStr;
			Logging.ToLog(subject);

			if (itemReport.Type == ReportsInfo.Type.TasksForItilium) {
				dataTableMainData = ExcelGeneral.ReadExcelFile(@"C:\Temp\Работы_январь_МИС.xlsx", "Лист1");
				return;
			}

			if (itemReport.Type == ReportsInfo.Type.MicroSipContactsBook) {
				dataTableMainData = MicroSipContactsBook.ReadContactsFile();
				return;
			}
			
			if (itemReport.Type == ReportsInfo.Type.RegistryMarks)
				dateBeginStr = "01.09.2018";

			if (itemReport.Type == ReportsInfo.Type.MESUsage ||
				itemReport.Type == ReportsInfo.Type.MESUsageFull) {
				int daysToLoad = (itemReport.DateEnd - itemReport.DateBegin).Days;
				List<DateTime> startDatesToLoad = new List<DateTime> {
					//itemReport.DateBegin.AddDays(-1 * daysToLoad - 1),  //For week comparison
					itemReport.DateBegin
				};

				for (int i = 0; i < startDatesToLoad.Count; i++) {
					string dateStartStr = startDatesToLoad[i].ToShortDateString();
					string dateEndStr = startDatesToLoad[i].AddDays(daysToLoad).ToShortDateString();
					Logging.ToLog("Получение данных из базы МИС Инфоклиника за период с " + 
						dateStartStr + " по " + dateEndStr);
					string period = (i + 1) + ". " + dateStartStr + "-" + dateEndStr;

					for (int y = 0; y <= daysToLoad; y++) {
						string dayToGetData = startDatesToLoad[i].AddDays(y).ToShortDateString();
						Logging.ToLog("Получение данных за день: " + dayToGetData);

						Dictionary<string, object> parametersMes = new Dictionary<string, object>() {
							{ "@dateBegin", dayToGetData },
							{ "@dateEnd", dayToGetData }
						};

						DataTable dataTablePart = dbClient.GetDataTable(itemReport.SqlQuery, parametersMes);
						Logging.ToLog("Получено строк: " + dataTablePart.Rows.Count);

						foreach (DataRow row in dataTablePart.Rows)
							row["PERIOD"] = period;

						if (dataTableMainData == null)
							dataTableMainData = dataTablePart;
						else
							dataTableMainData.Merge(dataTablePart);
					}
				}

				return;
			}

			if (itemReport.Type == ReportsInfo.Type.LicenseStatistics) {
				dataTableMainData = new DataTable();
				dataTableMainData.Columns.Add(new DataColumn("DB", typeof(string)));
				dataTableMainData.Columns.Add(new DataColumn("DATE", typeof(DateTime)));
				dataTableMainData.Columns.Add(new DataColumn("COUNT", typeof(int)));

				foreach (Tuple<string, string, string> item in infoclinicaDBs) {
					string dbName = item.Item1 + ":" + item.Item2 + "@" + item.Item3;
					Logging.ToLog("Получение данных из бд: " + dbName);
					try {
						dbClient = new FirebirdClient(
							item.Item1, 
							item.Item2, 
							Properties.Settings.Default.MisDbUser, 
							Properties.Settings.Default.MisDbPassword);

						DataTable dataTable = dbClient.GetDataTable(itemReport.SqlQuery, new Dictionary<string, object>());
						dataTableMainData.Rows.Add(new object[] {dbName, DateTime.Now, dataTable.Rows[0][0] });
					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
						dataTableMainData.Rows.Add(new object[] { dbName, DateTime.Now, -1 });
					}
				}

				return;
			}

			if (itemReport.Type == ReportsInfo.Type.LicenseEndingDates) {
				dataTableMainData = new DataTable();
				dataTableMainData.Columns.Add(new DataColumn("DB", typeof(string)));
				dataTableMainData.Columns.Add(new DataColumn("NOTBEFORE", typeof(DateTime)));

				foreach (Tuple<string, string, string> item in infoclinicaDBs) {
					string dbName = item.Item1 + ":" + item.Item2 + "@" + item.Item3;
					Logging.ToLog("Получение данных из бд: " + dbName);
					try {
						dbClient = new FirebirdClient(
							item.Item1,
							item.Item2,
							Properties.Settings.Default.MisDbUser,
							Properties.Settings.Default.MisDbPassword);

						DataTable dataTable = dbClient.GetDataTable(itemReport.SqlQuery, new Dictionary<string, object>());
						if (dataTable.Rows.Count > 0)
							dataTableMainData.Rows.Add(new object[] { dbName, DateTime.Parse(dataTable.Rows[0][0].ToString()) });
					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
						dataTableMainData.Rows.Add(new object[] { dbName, new DateTime() });
					}
				}

				return;
			}

			parameters = new Dictionary<string, object>() {
				{ "@dateBegin", dateBeginStr },
				{ "@dateEnd", dateEndStr }
			};

			if (itemReport.Type.ToString().StartsWith("TreatmentsDetails"))
				itemReport.SqlQuery = itemReport.SqlQuery.Replace("@jids", itemReport.JIDS);

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

						dataTableMainData = dbClient.GetDataTable(queryA8_2, parameters);
						dataTableWorkLoadA6 = dbClient.GetDataTable(queryA6, parameters);
						Logging.ToLog("Получено строк A6: " + dataTableWorkLoadA6.Rows.Count);
						dataTableWorkloadA11_10 = dbClient.GetDataTable(queryA11_10, parameters);
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

				dataTableUniqueServiceLab = dbClient.GetDataTable(sqlQueryUniqueServiceLab, parameters);

				Dictionary<string, object> parametersTotal = new Dictionary<string, object>() {
					{"@dateBegin",  DateTime.Parse("01.01." + itemReport.DateEnd.ToString("yyyy")).ToShortDateString() },
					{"@dateEnd", dateEndStr }
				};

				dataTableUniqueServiceTotal = dbClient.GetDataTable(itemReport.SqlQuery, parametersTotal);
				dataTableUniqueServiceLabTotal = dbClient.GetDataTable(sqlQueryUniqueServiceLab, parametersTotal);
			}

			if (itemReport.Type == ReportsInfo.Type.RegistryMotivation) {
				bool createNew = true;
				string fileToWrite = itemReport.TemplateFileName;

				foreach (KeyValuePair<string, string> query in registryMotivationQueries) {
					Logging.ToLog("Получение данных для листа: " + query.Key);
					DataTable dataTable = dbClient.GetDataTable(query.Value, parameters);
					Logging.ToLog("Получено строк: " + dataTable.Rows.Count);

					if (dataTable.Rows.Count > 0) {
						Logging.ToLog("Запись данных в Excel");
						itemReport.FileResult = ExcelGeneral.WriteDataTableToExcel(dataTable, subject, fileToWrite, query.Key, createNew);
						ExcelGeneral.CopyFormatting(itemReport.FileResult, query.Key, itemReport.Type);
						createNew = false;
						fileToWrite = itemReport.FileResult;
					}
				}
			}

			dataTableMainData = dbClient.GetDataTable(itemReport.SqlQuery, parameters);
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
					DataTable dataTablePriceExclusions = ExcelGeneral.ReadExcelFile(priceListToSiteSettingFilePath, sheetNameExclusions);
					Logging.ToLog("Считано строк: " + dataTablePriceExclusions.Rows.Count);
					DataTable dataTablePriceGrouping = ExcelGeneral.ReadExcelFile(priceListToSiteSettingFilePath, sheetNameGrouping);
					Logging.ToLog("Считано строк: " + dataTablePriceGrouping.Rows.Count);
					DataTable dataTablePricePriorities = ExcelGeneral.ReadExcelFile(priceListToSiteSettingFilePath, sheetNamePriorities);
					Logging.ToLog("Считано строк: " + dataTablePricePriorities.Rows.Count);

					dataTableMainData = PriceListToSite.PerformData(
						dataTableMainData, dataTablePriceExclusions, dataTablePriceGrouping, dataTablePricePriorities, out priceListToSiteEmptyFields);
				} catch (Exception e) {
					Logging.ToLog(e.StackTrace + Environment.NewLine + e.StackTrace);
					return;
				}
			}

            if (itemReport.Type == ReportsInfo.Type.FssInfo)
                FssInfo.PerformData(ref dataTableMainData);

			if (itemReport.Type == ReportsInfo.Type.AverageCheckRegular) {
				if (itemReport.DateBegin.Day == 1 &&
					itemReport.DateEnd.Day == DateTime.DaysInMonth(itemReport.DateBegin.Year, itemReport.DateBegin.Month) &&
					itemReport.DateBegin.Month == itemReport.DateEnd.Month &&
					itemReport.DateBegin.Year == itemReport.DateEnd.Year) {
					CultureInfo cultureInfoRU = CultureInfo.CreateSpecificCulture("ru");

					if (itemReport.DateBegin.Month == 1) {
						subjectAverageCheckPreviousWeek = ReportsInfo.AcceptedParameters[itemReport.Type] + ", месяц " +
							new DateTime(itemReport.DateBegin.Year - 1, 12, 1).ToString("MMMM", cultureInfoRU) +
							" год " + (itemReport.DateBegin.Year - 1) + " и месяц " +
							itemReport.DateBegin.ToString("MMMM", cultureInfoRU) + " год " + itemReport.DateBegin.Year;

						parametersAverageCheckPreviousWeek = new Dictionary<string, object> {
							{ "@dateBegin", new DateTime(itemReport.DateBegin.Year - 1, 12, 1).ToShortDateString() },
							{ "@dateEnd", new DateTime(itemReport.DateBegin.Year - 1, 12, 
								DateTime.DaysInMonth(itemReport.DateBegin.Year - 1, 12)).ToShortDateString() }
						};

					} else {
						subjectAverageCheckPreviousWeek = ReportsInfo.AcceptedParameters[itemReport.Type] + ", месяца " +
							new DateTime(itemReport.DateBegin.Year, itemReport.DateBegin.Month - 1, 1).ToString("MMMM", cultureInfoRU) +
							", " + itemReport.DateBegin.ToString("MMMM", cultureInfoRU) + " год " + itemReport.DateBegin.Year;

						parametersAverageCheckPreviousWeek = new Dictionary<string, object> {
							{ "@dateBegin", new DateTime(itemReport.DateBegin.Year, itemReport.DateBegin.Month - 1, 1).ToShortDateString() },
							{ "@dateEnd", new DateTime(itemReport.DateBegin.Year, itemReport.DateBegin.Month - 1, 
								DateTime.DaysInMonth(itemReport.DateBegin.Year, itemReport.DateBegin.Month - 1)).ToShortDateString() }
						};
					}

					subjectAverageCheckPreviousYear = ReportsInfo.AcceptedParameters[itemReport.Type] + ", месяц " + 
						itemReport.DateBegin.ToString("MMMM", cultureInfoRU) + " год " + 
						itemReport.DateBegin.Year + ", " + (itemReport.DateBegin.Year - 1);

					parametersAverageCheckPreviousYear = new Dictionary<string, object> {
						{ "@dateBegin", new DateTime(itemReport.DateBegin.Year - 1, itemReport.DateBegin.Month, 1).ToShortDateString() },
						{ "@dateEnd", new DateTime(itemReport.DateBegin.Year - 1, itemReport.DateBegin.Month, 
							DateTime.DaysInMonth(itemReport.DateBegin.Year - 1, itemReport.DateBegin.Month)).ToShortDateString() }
					};

				} else {
					#region previousWeek
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

					parametersAverageCheckPreviousWeek = new Dictionary<string, object> {
						{ "@dateBegin", dateBeginOriginal.Value.AddDays(
							-1 * (totalDays + 1)).ToShortDateString() },
						{ "@dateEnd", dateBeginOriginal.Value.AddDays(-1).ToShortDateString() }
					};
					#endregion

					#region previousYear
					DateTime previousYearWeekFirstDay = FirstDateOfWeekISO8601(dateBeginOriginal.Value.AddYears(-1).Year, reportWeekNumber);

					parametersAverageCheckPreviousYear = new Dictionary<string, object> {
						{ "@dateBegin", previousYearWeekFirstDay.ToShortDateString()},
						{ "@dateEnd", previousYearWeekFirstDay.AddDays(totalDays).ToShortDateString() }
					};
					#endregion
				}


				#region previousWeek
				Logging.ToLog("Получение данных из базы МИС Инфоклиника за период с " +
					parametersAverageCheckPreviousWeek["@dateBegin"] +
					" по " + parametersAverageCheckPreviousWeek["@dateEnd"]);
				dataTableAverageCheckPreviousWeek = dbClient.GetDataTable(
					itemReport.SqlQuery, parametersAverageCheckPreviousWeek);
				Logging.ToLog("Получено строк: " + dataTableAverageCheckPreviousWeek.Rows.Count);

				itemAverageCheckPreviousWeek = AverageCheck.PerformData(dataTableMainData, dataTableAverageCheckPreviousWeek);
				#endregion


				#region previousYear
				Logging.ToLog("Получение данных из базы МИС Инфоклиника за период с " +
					parametersAverageCheckPreviousYear["@dateBegin"] +
					" по " + parametersAverageCheckPreviousYear["@dateEnd"]);
				dataTableAverageCheckPreviousYear = dbClient.GetDataTable(
					itemReport.SqlQuery, parametersAverageCheckPreviousYear);
				Logging.ToLog("Получено строк: " + dataTableAverageCheckPreviousYear.Rows.Count);

				itemAverageCheckPreviousYear = AverageCheck.PerformData(dataTableMainData, dataTableAverageCheckPreviousYear);
				#endregion
			}

			if (itemReport.Type == ReportsInfo.Type.AverageCheckIGS)
				itemAverageCheckIGS = AverageCheck.PerformData(dataTableMainData);

			if (itemReport.Type == ReportsInfo.Type.CompetitiveGroups)
				ItemCompetitiveGroups = CompetitiveGroups.PerformData(dataTableMainData);

			if (itemReport.Type.ToString().StartsWith("TreatmentsDetails")) {
				TreatmentsDetails treatmentsDetails = new TreatmentsDetails();
				treatmentsDetails.PerformDataTable(dataTableMainData, itemReport.Type);
			}
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
					itemReport.Type == ReportsInfo.Type.FreeCellsWeek ||
					itemReport.Type == ReportsInfo.Type.FreeCellsMarketing) {
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
					itemReport.FileResult = MesUsage.WriteMesUsageTreatmentsToExcel(
						dataTableMainData, subject, itemReport.TemplateFileName);

				} else if (itemReport.Type == ReportsInfo.Type.MESUsageFull) {
					itemReport.FileResult = MesUsage.WriteMesUsageTreatmentsToExcel(
						dataTableMainData, subject, itemReport.TemplateFileName, true);

				} else if (itemReport.Type == ReportsInfo.Type.TelemedicineOnlyIngosstrakh) {
					itemReport.FileResult = ExcelGeneral.WriteDataTableToExcel(dataTableMainData,
														 subject,
														 itemReport.TemplateFileName,
														 type: itemReport.Type);

				} else if (itemReport.Type == ReportsInfo.Type.Workload) {
					for (int i = 0; i < workloadResultFiles.Count; i++) {
						string key = workloadResultFiles.Keys.ElementAt(i);
						Logging.ToLog("Филиал: " + key);

						workloadResultFiles[key] = ExcelGeneral.WriteDataTableToExcel(dataTableWorkLoadA6,
															 subject + " " + key,
															 itemReport.TemplateFileName,
															 "Услуги Мет. 1",
															 true,
															 key);

						if (string.IsNullOrEmpty(workloadResultFiles[key]))
							continue;

						ExcelGeneral.WriteDataTableToExcel(dataTableWorkloadA11_10,
												subject,
												workloadResultFiles[key],
												"Искл. услуги",
												false,
												key);

						ExcelGeneral.WriteDataTableToExcel(dataTableMainData,
												subject,
												workloadResultFiles[key],
												"Расчет",
												false,
												key);
					}

				} else if (itemReport.Type == ReportsInfo.Type.Robocalls) {
					itemReport.FileResult = ExcelGeneral.WriteDataTableToTextFile(dataTableMainData,
															subject,
															itemReport.TemplateFileName);

				} else if (itemReport.Type == ReportsInfo.Type.PriceListToSite) {
					itemReport.FileResult = ExcelGeneral.WriteDataTableToExcel(
						dataTableMainData,
						subject,
						itemReport.TemplateFileName,
						type: itemReport.Type);

					fileToUpload = ExcelGeneral.WriteDataTableToTextFile(
						dataTableMainData,
						"BzPriceListToUpload",
						saveAsJson: true);

				} else if (itemReport.Type == ReportsInfo.Type.TimetableToProdoctorovRu) {
					fileToUpload = TimetableToProdoctorovRu.PerformData(dataTableMainData);

				} else if (itemReport.Type == ReportsInfo.Type.UniqueServices ||
					itemReport.Type == ReportsInfo.Type.UniqueServicesRegions) {
					itemReport.FileResult = UniqueServices.Process(dataTableMainData,
														 dataTableUniqueServiceTotal,
														 dataTableUniqueServiceLab,
														 dataTableUniqueServiceLabTotal,
														 subject,
														 itemReport.TemplateFileName,
														 dateBeginStr + " - " + dateEndStr,
														 itemReport.Type);

				} else if (itemReport.Type == ReportsInfo.Type.AverageCheckRegular) {
					itemReport.FileResult =
						AverageCheck.WriteAverageCheckToExcel(itemAverageCheckPreviousWeek,
							subjectAverageCheckPreviousWeek, itemReport.TemplateFileName);
					itemReport.FileResultAverageCheckPreviousYear =
						AverageCheck.WriteAverageCheckToExcel(itemAverageCheckPreviousYear,
							subjectAverageCheckPreviousYear, itemReport.TemplateFileName);

				} else if (itemReport.Type == ReportsInfo.Type.AverageCheckIGS) {
					itemReport.FileResult = AverageCheck.WriteAverageCheckToExcel(itemAverageCheckIGS, subject, itemReport.TemplateFileName);

				} else if (itemReport.Type == ReportsInfo.Type.CompetitiveGroups) {
					itemReport.FileResult =
						CompetitiveGroups.WriteAverageCheckToExcel(
							ItemCompetitiveGroups, subject, itemReport.TemplateFileName);

				} else if (itemReport.Type == ReportsInfo.Type.TimetableToSite) {
					fileToUpload = TimetableToSite.ParseAndWriteToJson(dataTableMainData);
					itemReport.FileResult = fileToUpload;

				} else if (itemReport.Type == ReportsInfo.Type.MicroSipContactsBook) {
					itemReport.FileResult = MicroSipContactsBook.WriteToFile(dataTableMainData);

				} else if (itemReport.Type == ReportsInfo.Type.TasksForItilium) {
					itemReport.FileResult = TasksForItilium.SendTasks(dataTableMainData);

				} else if (itemReport.Type == ReportsInfo.Type.RegistryMotivation &&
					!string.IsNullOrEmpty(itemReport.FileResult)) {
					ExcelGeneral.WriteDataTableToExcel(dataTableMainData, subject, itemReport.FileResult, "Данные", false);

				} else if (itemReport.Type == ReportsInfo.Type.LicenseEndingDates) {
					foreach (DataRow row in dataTableMainData.Rows) {
						DateTime dateTime = (DateTime)row[1];
						int daysLeft = (int)(dateTime - DateTime.Now).TotalDays;
						string message = string.Empty;

						string db = row[0].ToString();
						if (daysLeft <= 7)
							message += Environment.NewLine + "DB: " + db + " до окончания лицензии осталось дней: " + daysLeft;
						else
							Logging.ToLog("DB: " + db + " до окончания лицензии осталось дней: " + daysLeft);

						if (!string.IsNullOrEmpty(message)) 
							SystemMail.SendMail(subject, "На отдел сопровождения МИС: " + Environment.NewLine + message, itemReport.MailTo);
					}

					return;

				} else {
					itemReport.FileResult = ExcelGeneral.WriteDataTableToExcel(dataTableMainData,
														 subject,
														 itemReport.TemplateFileName,
														 type: itemReport.Type);
				}

				if (File.Exists(itemReport.FileResult) ||
					itemReport.Type == ReportsInfo.Type.Workload ||
					itemReport.Type == ReportsInfo.Type.TasksForItilium) {
					bool isPostProcessingOk = true;

					switch (itemReport.Type) {
						case ReportsInfo.Type.FreeCellsDay:
						case ReportsInfo.Type.FreeCellsWeek:
							isPostProcessingOk = FreeCells.Process(itemReport.FileResult, dateBeginOriginal.Value, itemReport.DateEnd);
							break;

						case ReportsInfo.Type.UnclosedProtocolsWeek:
						case ReportsInfo.Type.UnclosedProtocolsMonth:
							isPostProcessingOk = UnclosedProtocols.Process(itemReport.FileResult);
							break;

						case ReportsInfo.Type.MESUsage:
							isPostProcessingOk = MesUsage.Process(itemReport.FileResult);
							break;

						case ReportsInfo.Type.MESUsageFull:
							isPostProcessingOk = MesUsage.Process(itemReport.FileResult, true);
							break;

						case ReportsInfo.Type.OnlineAccountsUsage:
							isPostProcessingOk = OnlineAccounts.Process(itemReport.FileResult);
							break;

						case ReportsInfo.Type.TelemedicineOnlyIngosstrakh:
						case ReportsInfo.Type.TelemedicineAll:
							isPostProcessingOk = Telemedicine.Process(itemReport.FileResult);
							break;

						case ReportsInfo.Type.NonAppearance:
							isPostProcessingOk = NonAppearance.Process(itemReport.FileResult, dataTableMainData);
							break;

						case ReportsInfo.Type.VIP_MSSU:
						case ReportsInfo.Type.VIP_Moscow:
						case ReportsInfo.Type.VIP_MSKM:
						case ReportsInfo.Type.VIP_PND:
							isPostProcessingOk = VIP.Process(itemReport.FileResult, itemReport.PreviousFile);
							break;

						case ReportsInfo.Type.RegistryMarks:
							isPostProcessingOk = RegistryMarks.Process(
								itemReport.FileResult, dataTableMainData, dateBeginOriginal.Value);
							break;

						case ReportsInfo.Type.Workload:
							bool isAllOk = true;
							Logging.ToLog("Пост-обработка");
							foreach (string currentFileResult in workloadResultFiles.Values) {
								Logging.ToLog("Файл: " + currentFileResult);

								if (string.IsNullOrEmpty(currentFileResult))
									continue;

								if (!Workload.Process(currentFileResult))
									isAllOk = false;
							}

							isPostProcessingOk = isAllOk;
							break;

                        case ReportsInfo.Type.PriceListToSite:
                            isPostProcessingOk = PriceListToSite.Process(itemReport.FileResult);
                            break;

                        case ReportsInfo.Type.GBooking:
						case ReportsInfo.Type.PersonalAccountSchedule:
						case ReportsInfo.Type.ProtocolViewCDBSyncEvent:
							isPostProcessingOk = ExcelGeneral.CopyFormatting(itemReport.FileResult);
							break;

                        case ReportsInfo.Type.FssInfo:
                            isPostProcessingOk = FssInfo.Process(itemReport.FileResult);
                            break;

                        case ReportsInfo.Type.RecordsFromInsuranceCompanies:
                            isPostProcessingOk = RecordsFromInsuranceCompanies.Process(itemReport.FileResult);
                            break;

						case ReportsInfo.Type.AverageCheckRegular:
							isPostProcessingOk = AverageCheck.Process(
								itemReport.FileResult, parameters, parametersAverageCheckPreviousWeek);
							isPostProcessingOk &= AverageCheck.Process(
								itemReport.FileResultAverageCheckPreviousYear, parameters, parametersAverageCheckPreviousYear);
							break;

						case ReportsInfo.Type.AverageCheckIGS:
							isPostProcessingOk = AverageCheck.Process(
								itemReport.FileResult, parameters, null);
							break;

						case ReportsInfo.Type.CompetitiveGroups:
							isPostProcessingOk = CompetitiveGroups.Process(itemReport.FileResult, parameters);
							break;

						case ReportsInfo.Type.FirstTimeVisitPatients:
							isPostProcessingOk = FirstTimeVisitPatients.Process(itemReport.FileResult, dataTableMainData);
							break;

						case ReportsInfo.Type.FreeCellsMarketing:
							isPostProcessingOk = FreeCells.Process(itemReport.FileResult, dateBeginOriginal.Value, itemReport.DateEnd, true);
							break;

						case ReportsInfo.Type.EmergencyCallsQuantity:
							isPostProcessingOk = ExcelGeneral.CopyFormatting(itemReport.FileResult);
							break;

						case ReportsInfo.Type.RegistryMotivation:
							isPostProcessingOk = RegistryMotivation.Process(itemReport.FileResult);
							break;


						case ReportsInfo.Type.Promo:
							isPostProcessingOk = Promo.Process(itemReport.FileResult);
							break;

						default:
							break;
					}

					if (itemReport.Type.ToString().StartsWith("TreatmentsDetails"))
						isPostProcessingOk = ExcelGeneral.CopyFormatting(itemReport.FileResult);

					if (isPostProcessingOk) {
						body = "Отчет во вложении";
						Logging.ToLog("Данные сохранены в файл: " + (itemReport.Type == ReportsInfo.Type.Workload ?
							string.Join("; ", workloadResultFiles.Values) :
							itemReport.FileResult));

						if (itemReport.Type == ReportsInfo.Type.TasksForItilium)
							body = itemReport.FileResult;
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
			if (File.Exists(destFile))
				try {
					File.Delete(destFile);
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}

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
				url = "https://old.klinikabudzdorov.ru/export/price/file_input.php";
				method = WebRequestMethods.Http.Post;

			} else if (itemReport.Type == ReportsInfo.Type.TimetableToProdoctorovRu) {
				PostDataToServer();
				return;

			} else if (itemReport.Type == ReportsInfo.Type.TimetableToSite) {
				url = "https://old.klinikabudzdorov.ru/export/schedule/file_input.php";
				method = WebRequestMethods.Http.Post;

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

					if (itemReport.Type == ReportsInfo.Type.TimetableToSite)
						body = response;

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
	}
}
