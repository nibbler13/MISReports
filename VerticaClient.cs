﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Vertica.Data.VerticaClient;

namespace MISReports {
	class VerticaClient : IDbClient {
		private VerticaConnection connection;

		public VerticaClient(string host, string database, string user, string password) {
			VerticaConnectionStringBuilder builder = new VerticaConnectionStringBuilder {
				Host = host,
				Database = database,
				User = user,
				Password = password
			};

			connection = new VerticaConnection(builder.ToString());
			IsConnectionOpened();
		}

		private bool IsConnectionOpened() {
			if (connection.State != ConnectionState.Open) {
				try {
					connection.Open();
				} catch (Exception e) {
					string subject = (Program.itemReport is null ? string.Empty : Program.itemReport.Type.ToString()) + " Ошибка подключения к БД";
					string body = e.Message + Environment.NewLine + e.StackTrace;
					SystemMail.SendMail(subject, body, Properties.Settings.Default.MailCopy);
					Logging.ToLog(subject + " " + body);
					Program.hasError = true;
				}
			}

			return connection.State == ConnectionState.Open;
		}

		public DataTable GetDataTable(string query, Dictionary<string, object> parameters = null) {
			DataTable dataTable = new DataTable();

			if (!IsConnectionOpened())
				return dataTable;

			try {
				using (VerticaCommand command = new VerticaCommand(query, connection)) {
					if (parameters != null && parameters.Count > 0)
						foreach (KeyValuePair<string, object> parameter in parameters)
							if (query.Contains(parameter.Key))
								command.Parameters.Add(new VerticaParameter(parameter.Key, parameter.Value));

					using (VerticaDataAdapter fbDataAdapter = new VerticaDataAdapter(command))
						fbDataAdapter.Fill(dataTable);
				}
			} catch (Exception e) {
				string subject = (Program.itemReport is null ? string.Empty : Program.itemReport.Type.ToString()) + " Ошибка выполнения запроса к БД";
				string body = e.Message + Environment.NewLine + e.StackTrace;
				SystemMail.SendMail(subject, body, Properties.Settings.Default.MailCopy);
				Logging.ToLog(subject + " " + body);
				connection.Close();
				Program.hasError = true;
			}

			return dataTable;
		}

        public bool ExecuteUpdateQuery(string query, Dictionary<string, object> parameters) {
			bool updated = false;

			if (!IsConnectionOpened())
				return updated;

			try {
				VerticaCommand update = new VerticaCommand(query, connection);

				if (parameters.Count > 0) {
					foreach (KeyValuePair<string, object> parameter in parameters)
						update.Parameters.Add(new VerticaParameter(parameter.Key, parameter.Value));
				}

				updated = update.ExecuteNonQuery() > 0 ? true : false;
			} catch (Exception e) {
				string subject = "Ошибка выполнения запроса к БД";
				string body = e.Message + Environment.NewLine + e.StackTrace;
				SystemMail.SendMail(subject, body, Properties.Settings.Default.MailCopy);
				Logging.ToLog(subject + " " + body);
				connection.Close();
			}

			return updated;
		}

        public void Close() {
			connection.Close();
        }

        public string GetName() {
			return connection.Database;
        }
    }
}
