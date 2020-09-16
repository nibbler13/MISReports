using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace MISReports.DataHandlers {
	public class PatientToSha1 {
		public static DataTable PerformDataTable(DataTable dataTable) {
			DataRow dataRowIvanov = dataTable.NewRow();
			dataRowIvanov["PCODE"] = "0";
			dataRowIvanov["FULLNAME"] = "Иванов Иван Иванович";
			dataRowIvanov["BDATE"] = "01.01.1980";
			dataRowIvanov["PASPSER"] = "56 31";
			dataRowIvanov["PASPNUM"] = "123456";
			dataRowIvanov["MAX"] = "01.01.2010";
			dataTable.Rows.InsertAt(dataRowIvanov, 0);


			DataTable dataTableResult = new DataTable();
			dataTableResult.Clear();
			dataTableResult.Columns.Add("SOURCE");
			dataTableResult.Columns.Add("CLIENTID");
			dataTableResult.Columns.Add("HASH1");
			dataTableResult.Columns.Add("HASH2");
			dataTableResult.Columns.Add("STATE");

			string source = "BZ";

			foreach (DataRow row in dataTable.Rows) {
				string pcode = row["PCODE"].ToString();
				string fullname = row["FULLNAME"].ToString();
				string bdate = row["BDATE"].ToString().Replace(" 0:00:00", "");
				string paspser = row["PASPSER"].ToString();
				string paspnum = row["PASPNUM"].ToString();
				string max = row["MAX"].ToString().Replace(" 0:00:00", "");

				string hash1Original = fullname + bdate + paspser + paspnum;
				string hash2Original = fullname + bdate;

				string hash1 = GetSha1Hash(GetNormalizedString(hash1Original));
				string hash2 = GetSha1Hash(GetNormalizedString(hash2Original));

				int state = 0;
				if (!string.IsNullOrEmpty(max)) {
					DateTime maxDate = DateTime.Parse(max);
					if (maxDate.Date >= DateTime.Now)
						state = 2;
					else if ((DateTime.Now - maxDate).TotalDays / 365 <= 5)
						state = 1;
				}

				DataRow dataRow = dataTableResult.NewRow();
				dataRow["SOURCE"] = source;
				dataRow["CLIENTID"] = pcode;
				dataRow["HASH1"] = hash1 + GetNormalizedString(hash1Original).Substring(0, 3).ToUpper();
				dataRow["HASH2"] = hash2 + GetNormalizedString(hash2Original).Substring(1, 3).ToUpper();
				dataRow["STATE"] = state;
				dataTableResult.Rows.Add(dataRow);
			}


			return dataTableResult;
		}

		private static string GetNormalizedString(string text) {
			return text.Replace(" ", "").Replace(".", "").ToLower();
		}

		private static string GetSha1Hash(string text) {
			using (SHA1 sha1 = SHA1.Create()) {
				byte[] sourceBytes = Encoding.UTF8.GetBytes(text);
				byte[] hashBytes = sha1.ComputeHash(sourceBytes);
				string hash = BitConverter.ToString(hashBytes).Replace("-", string.Empty);
				return hash;
			}
		}
	}
}
