using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MISReports.ExcelHandlers {
	static class TimetableToSite {
		public static string ParseAndWriteToJson(DataTable dataTable) {
			List<ItemRow> rows = new List<ItemRow>();

			try {
				foreach (DataRow dataRow in dataTable.Rows) {
					ItemRow row = new ItemRow {
						FILIAL = dataRow["FILIAL"].ToString(),
						CASHID = dataRow["CASHID"].ToString(),
						CASHNAME = dataRow["CASHNAME"].ToString(),
						SHORTNAME = dataRow["SHORTNAME"].ToString(),
						FULLNAME = dataRow["FULLNAME"].ToString(),
						DEPNUM = dataRow["DEPNUM"].ToString(),
						DEPNAME = dataRow["DEPNAME"].ToString(),
						DCODE = dataRow["DCODE"].ToString(),
						DNAME = dataRow["DNAME"].ToString(),
						DOCTPOST_ID = dataRow["DOCTPOST_ID"].ToString(),
						DOCTPOST = dataRow["DOCTPOST"].ToString(),
						WDATE = dataRow["WDATE"].ToString(),
						STARTTIME = dataRow["STARTTIME"].ToString(),
						ENDTIME = dataRow["ENDTIME"].ToString(),
						INTERVALTIME = dataRow["INTERVALTIME"].ToString()
					};

					rows.Add(row);
				}
			} catch (Exception) { }

			string json = Newtonsoft.Json.JsonConvert.SerializeObject(rows, Newtonsoft.Json.Formatting.Indented);
			string resultFile = ExcelGeneral.GetResultFilePath("TimetableToSite", isPlainText: true).Replace(".txt", ".json");
			try {
				File.WriteAllText(resultFile, json);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			return resultFile;
		}

		private class ItemRow {
			public string FILIAL { get; set; } //6
			public string CASHID { get; set; }
			public string CASHNAME { get; set; }
			public string SHORTNAME { get; set; } //КУТУЗ
			public string FULLNAME { get; set; } //КУТУЗ ООО "Клиника ЛМС", 119146 Москва, Комсомольский проспект, 28, тел (495) 782 88 82, e-mail: info @bzklinika.ru
			public string DEPNUM { get; set; } //10029098
			public string DEPNAME { get; set; } //ПОМОЩЬ НА ДОМУ
			public string DCODE { get; set; } //60000445
			public string DNAME { get; set; } //Терехов К. А.
			public string DOCTPOST_ID { get; set; } //990000098
			public string DOCTPOST { get; set; } //Терапевт (ПНД)
			public string WDATE { get; set; } //03.12.2019
			public string STARTTIME { get; set; } //08:00
			public string ENDTIME { get; set; } //20:00
			public string INTERVALTIME { get; set; } //08:00 - 20:00
		}
	}
}
