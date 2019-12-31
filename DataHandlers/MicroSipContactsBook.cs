using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace MISReports.ExcelHandlers {
	class MicroSipContactsBook : ExcelGeneral {
		public static DataTable ReadContactsFile() {
			DataTable dataTable = new DataTable();
			dataTable.Columns.Add(new DataColumn("name", typeof(string)));
			dataTable.Columns.Add(new DataColumn("phoneNumber", typeof(string)));

			string filePath = @"\\budzdorov.ru\NETLOGON\телефоны клиники Будь Здоров.xls";
			List<string> sheetNames = ReadSheetNames(filePath);
			foreach (string sheetName in sheetNames) {
				string sheetNameCleared = sheetName.TrimStart('\'').TrimEnd('\'');
				DataTable dataTableSheet = ReadExcelFile(filePath, sheetNameCleared);

				foreach (DataRow dataRow in dataTableSheet.Rows) {
					string phoneNumber = dataRow["F4"].ToString();
					if (string.IsNullOrEmpty(phoneNumber))
						continue;

					phoneNumber = phoneNumber.Replace("-", "");

					if (phoneNumber.ToLower().Equals("внутренний номер"))
						continue;

					string name = dataRow["F3"].ToString();
					if (!string.IsNullOrEmpty(name))
						name += ", ";

					string position = dataRow["F2"].ToString();
					//if (!string.IsNullOrEmpty(position))
					//	position += ", ";

					DataRow dataRowInfo = dataTable.NewRow();
					dataRowInfo["name"] = sheetNameCleared.Replace("$", "") + ", " + name + position ;
					dataRowInfo["phoneNumber"] = phoneNumber;
					dataTable.Rows.Add(dataRowInfo);
				}
			}

			return dataTable;
		}

		public static string WriteToFile(DataTable dataTable) {
			string resultFile = string.Empty;

			StringBuilder stringBuilder = new StringBuilder();
			stringBuilder.AppendLine("<?xml version=\"1.0\"?>");
			stringBuilder.AppendLine("<contacts>");

			foreach (DataRow dataRow in dataTable.Rows) {
				string name = dataRow["name"].ToString();
				string phoneNumber = dataRow["phoneNumber"].ToString();
				stringBuilder.AppendLine($"<contact number=\"{phoneNumber}\"  name=\"{name}\"  presence=\"0\"  directory=\"0\" >");
				stringBuilder.AppendLine("</contact>");
			}
					   
			stringBuilder.AppendLine("</contacts>");
			stringBuilder.AppendLine("");
			string resultString = stringBuilder.ToString();

			resultFile = GetResultFilePath("Contacts", isPlainText: true);
			resultFile = Path.Combine(Path.GetDirectoryName(resultFile), "Contacts.xml");

			if (File.Exists(resultFile))
				File.Delete(resultFile);

			File.WriteAllText(resultFile, resultString);

			return resultFile;
		}
	}
}
