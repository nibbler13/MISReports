using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports.ExcelHandlers {
	class TasksForItilium {
		public static string SendTasks(DataTable dataTable) {
			DateTime dateNow = DateTime.Now;
			string recipient = "stp@bzklinika.ru";
			int tasksSended = 0;

			foreach (DataRow dataRow in dataTable.Rows) {
				try {
					string date = dataRow[0].ToString();
					if (!DateTime.TryParse(date, out DateTime dateTime))
						continue;

					if (!dateTime.Date.Equals(dateNow.Date))
						continue;

					string task = dataRow[1].ToString();
					string responsible = dataRow[2].ToString();
					string initiator = dataRow[3].ToString();

					if (string.IsNullOrEmpty(task) ||
						string.IsNullOrEmpty(responsible) ||
						string.IsNullOrEmpty(initiator)) {
						Logging.ToLog("Строка содержит пустые ячейки: задача - " + task + 
							", ответственный - " + responsible + ", инициатор - " + initiator);
						Logging.ToLog("Пропуск");
						continue;
					}

					string body = "Назначить на: " + responsible + Environment.NewLine +
						"Задача: " + task + Environment.NewLine +
						"Инициатор: " + initiator;

					SystemMail.SendMail("Заявка", body, recipient, addSignature: false);
					tasksSended++;
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			}

			return "Задач отправлено: " + tasksSended;
		}
	}
}
