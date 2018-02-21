using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports {
	public class MESUsage {
		public static void CreateAndSendReport(FirebirdClient firebirdClient, string dateBegin, string dateEnd) {
			string subject = "Отчет по использованию МЭС с " + dateBegin + " по " + dateEnd;
			SystemLogging.LogMessageToFile(subject);

			Dictionary<string, object> parameters = new Dictionary<string, object>() {
					{ "@dateBegin", dateBegin },
					{ "@dateEnd", dateEnd } };

			SystemLogging.LogMessageToFile("Получение данных из базы МИС Инфоклиника за период с " + dateBegin + " по " + dateEnd);
			DataTable dataTable = firebirdClient.GetDataTable(
				Properties.Settings.Default.MisDbSqlGetMESUsage, parameters);

			SystemLogging.LogMessageToFile("Получено строк: " + dataTable.Rows.Count);

			Dictionary<string, ItemTreatment> treatments = new Dictionary<string, ItemTreatment>();
			foreach (DataRow row in dataTable.Rows) {
				try {
					string treatcode = row["TREATCODE"].ToString();
					string mid = row["MID"].ToString();
					string listMES = row["LISTMES"].ToString();
					string listReferrals = row["LISTREFERRALS"].ToString();
					string[] arrayMES = new string[0];
					string[] arrayReferrals = new string[0];
					if (!string.IsNullOrEmpty(listMES))
						arrayMES = listMES.Split(';');
					if (!string.IsNullOrEmpty(listReferrals))
						arrayReferrals = listReferrals.Split(';');

					if (treatments.ContainsKey(treatcode)) {
						treatments[treatcode].ListMES.AddRange(arrayMES);

						if (string.IsNullOrEmpty(mid))
							treatments[treatcode].ListReferralsFromDoc.AddRange(arrayReferrals);
						else
							treatments[treatcode].ListReferralsFromMes.AddRange(arrayReferrals);
					} else {
						ItemTreatment treatment = new ItemTreatment() {
							TREATDATE = row["TREATDATE"].ToString(),
							CLIENTNAME = row["CLIENTNAME"].ToString(),
							HISTNUM = row["HISTNUM"].ToString(),
							DOCNAME = row["DOCNAME"].ToString(),
							DOCPOSITION = row["DOCPOSITION"].ToString(),
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

						treatments.Add(treatcode, treatment);
					}
				} catch (Exception e) {
					SystemLogging.LogMessageToFile(e.Message);
				}
			}

			SystemLogging.LogMessageToFile("Запись данных в файл Excel");
			string fileResult = NpoiExcelMESUsage.WriteTreatmentsToExcel(treatments, subject);

			string mailTo = Properties.Settings.Default.MailTo;
			string body = "Отчет во вложении";

			if (File.Exists(fileResult)) {
				SystemLogging.LogMessageToFile("Данные сохранены в файл: " + fileResult);

				if (dataTable.Rows.Count > 0) {
					body = "Отчет во вложении";
				} else {
					body = "Отсутствуют данные за период " + dateBegin + "-" + dateEnd;
					fileResult = string.Empty;
				}
			} else {
				mailTo = Properties.Settings.Default.MailCopy;
				body = "Не удалось записать данные в файл: " + fileResult;
				SystemLogging.LogMessageToFile(body);
				fileResult = "";
			}

			//SystemMail.SendMail(subject, body, mailTo, fileResult);
		}

		public class ItemTreatment {
			public string TREATDATE { get; set; } = string.Empty;
			public string FILIAL { get; set; } = string.Empty;
			public string DEPNAME { get; set; } = string.Empty;
			public string DOCNAME { get; set; } = string.Empty;
			public string DOCPOSITION { get; set; } = string.Empty;
			public string HISTNUM { get; set; } = string.Empty;
			public string CLIENTNAME { get; set; } = string.Empty;
			public string MKBCODE { get; set; } = string.Empty;
			public string AGE { get; set; } = string.Empty;
			public List<string> ListMES { get; set; } = new List<string>();
			public List<string> ListReferralsFromMes { get; set; } = new List<string>();
			public List<string> ListReferralsFromDoc { get; set; } = new List<string>();
		}
	}
}
