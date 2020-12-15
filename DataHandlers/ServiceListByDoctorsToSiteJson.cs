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
	class ServiceListByDoctorsToSiteJson {
		public static string ParseDataTableToJsonAndWriteToFile(DataTable dataTable) {
			string resultFile = string.Empty;

			if (dataTable == null)
				return resultFile;

			Dictionary<long, List<string>> doctorList = new Dictionary<long, List<string>>();

			try {
				foreach (DataRow row in dataTable.Rows) {
					long dcode = (long)row["dcode"];
					string kodoper = row["wschema_kodoper"].ToString();

					if (!doctorList.ContainsKey(dcode))
						doctorList.Add(dcode, new List<string>());

					doctorList[dcode].Add(kodoper);
				}
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			string json = JsonConvert.SerializeObject(doctorList);
			resultFile = ExcelGeneral.GetResultFilePath("ServiceListByDoctorsToSiteJson", isPlainText: true).Replace(".txt", ".json");
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
