﻿using System;
using System.Data;
using System.Collections.Generic;
using FirebirdSql.Data.FirebirdClient;
using System.Windows;

namespace MISReports {
	public class FirebirdClient : IDbClient {
		private FbConnection connection;

		public FirebirdClient(string ipAddress, string baseName, string user, string pass) {
			FbConnectionStringBuilder cs = new FbConnectionStringBuilder {
				DataSource = ipAddress,
				Database = baseName,
				UserID = user,
				Password = pass,
				Charset = "NONE",
				Pooling = false
			};

			connection = new FbConnection(cs.ToString());
			IsConnectionOpened();
		}

		public void Close() {
			connection.Close();
		}

		private bool IsConnectionOpened() {
			if (connection.State != ConnectionState.Open) {
				try {
					connection.Open();
				} catch (Exception e) {
					string subject = "Ошибка подключения к БД";
					string body = e.Message + Environment.NewLine + e.StackTrace;
					SystemMail.SendMail(subject, body, Properties.Settings.Default.MailCopy);
					Logging.ToLog(subject + " " + body);
				}
			}

			return connection.State == ConnectionState.Open;
		}

		public DataTable GetDataTable(string query, Dictionary<string, object> parameters) {
			DataTable dataTable = new DataTable();

			if (!IsConnectionOpened())
				return dataTable;
			
			try {
				using (FbCommand command = new FbCommand(query, connection)) { 
					if (parameters.Count > 0) 
						foreach (KeyValuePair<string, object> parameter in parameters)
							command.Parameters.AddWithValue(parameter.Key, parameter.Value);
				
					using (FbDataAdapter fbDataAdapter = new FbDataAdapter(command)) 
						fbDataAdapter.Fill(dataTable);
				}
			} catch (Exception e) {
				string subject = (Program.itemReport is null ? string.Empty : Program.itemReport.Type.ToString()) +  " Ошибка выполнения запроса к БД";
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
				FbCommand update = new FbCommand(query, connection);

				if (parameters.Count > 0) {
					foreach (KeyValuePair<string, object> parameter in parameters)
						update.Parameters.AddWithValue(parameter.Key, parameter.Value);
				}

				updated = update.ExecuteNonQuery() > 0 ? true : false;
			} catch (Exception e) {
				string subject = (Program.itemReport is null ? string.Empty : Program.itemReport.Type.ToString()) + " Ошибка выполнения запроса к БД";
				string body = e.Message + Environment.NewLine + e.StackTrace;
				SystemMail.SendMail(subject, body, Properties.Settings.Default.MailCopy);
				Logging.ToLog(subject + " " + body);
				connection.Close();
				Program.hasError = true;
			}

			return updated;
		}

        public string GetName() {
			return connection.Database;
        }
    }
}
