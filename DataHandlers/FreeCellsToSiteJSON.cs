using MISReports.ExcelHandlers;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports.DataHandlers {
	class FreeCellsToSiteJSON {
		public static string ParseDataTableToJsonAndWriteToFile(DataTable dataTable) {
			string resultFile = string.Empty;

			Dictionary<long, ItemDoctor> doctorList = new Dictionary<long, ItemDoctor>();
			if (dataTable == null)
				return resultFile;

			try {
				foreach (DataRow row in dataTable.Rows) {
					long dcode = (long)row["dcode"];
					string date = ((DateTime)row["date"]).ToShortDateString();

					if (!doctorList.ContainsKey(dcode))
						doctorList.Add(dcode, new ItemDoctor { DCode = dcode });

					if (!doctorList[dcode].DaysCheck.ContainsKey(date))
						doctorList[dcode].DaysCheck.Add(date, new Dictionary<string, ItemInterval>());

					if (row["start"] == null || string.IsNullOrEmpty(row["start"].ToString()) ||
						row["end"] == null || string.IsNullOrEmpty(row["end"].ToString()) ||
						row["status"] == null || string.IsNullOrEmpty(row["status"].ToString()))
						continue;

					TimeSpan start = (TimeSpan)row["start"];
					TimeSpan end = (TimeSpan)row["end"];
					string startTime = start.ToString("hh\\:mm");
					string endTime = end.ToString("hh\\:mm");
					string status = (string)row["status"];
					int statusResult = status.Equals("Свободное время") || status.Equals("Резерв с возможностью записи") ? 1 : 2;

					ItemInterval interval = new ItemInterval {
						TStart = startTime,
						TEnd = endTime,
						Status = statusResult
					};

					string intervalKey = startTime + "-" + endTime;

					if (!doctorList[dcode].DaysCheck[date].ContainsKey(intervalKey))
						doctorList[dcode].DaysCheck[date].Add(intervalKey, interval);
					else {
						if (doctorList[dcode].DaysCheck[date][intervalKey].Status == 1)
							if (interval.Status == 1)
								doctorList[dcode].DaysCheck[date][intervalKey] = interval;
					}
				}
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			foreach (ItemDoctor doc in doctorList.Values) 
				foreach (string date in doc.DaysCheck.Keys) 
					doctorList[doc.DCode].Days.Add(date, doc.DaysCheck[date].Values.ToList());

			string json = JsonConvert.SerializeObject(doctorList.Values); 
			resultFile = ExcelGeneral.GetResultFilePath("FreeCellsToSiteJSON", isPlainText: true).Replace(".txt", ".json");
			try {
				File.WriteAllText(resultFile, json);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			return resultFile;
		}

		public class ItemDoctor {
			public long DCode { get; set; }
			public Dictionary<string, List<ItemInterval>> Days { get; set; } = new Dictionary<string, List<ItemInterval>>();

			[JsonIgnore]
			public Dictionary<string, Dictionary<string, ItemInterval>> DaysCheck { get; set; } =
				new Dictionary<string, Dictionary<string, ItemInterval>>();
		}

		public class ItemInterval {
			public string TStart { get; set; }
			public string TEnd { get; set; }
			public int Status { get; set; } //1 - Свободно, 2 - занято
		}
	}
}
