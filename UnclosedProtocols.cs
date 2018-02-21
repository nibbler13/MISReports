using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports {
	class UnclosedProtocols {
		public static void CreateAndSendReport(FirebirdClient firebirdClient, string dateBegin, string dateEnd) {
			string subject = "Отчет по неподписанным протоколам с " + dateBegin + " по " + dateEnd;
			SystemLogging.LogMessageToFile(subject);

			Dictionary<string, object> parameters = new Dictionary<string, object>() {
					{ "@dateBegin", dateBegin },
					{ "@dateEnd", dateEnd } };

			SystemLogging.LogMessageToFile("Получение данных из базы МИС Инфоклиника за период с " + dateBegin + " по " + dateEnd);
			DataTable dataTable = firebirdClient.GetDataTable(
				Properties.Settings.Default.MisDbSqlGetUnclosedProtocols, parameters);

			SystemLogging.LogMessageToFile("Получено строк: " + dataTable.Rows.Count);

			SystemLogging.LogMessageToFile("Запись данных в файл Excel");
			string fileResult = NpoiExcelUnclosedProtocols.WriteDataTableToExcel(dataTable, subject);

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

			SystemMail.SendMail(subject, body, mailTo, fileResult);
		}
	}
}
