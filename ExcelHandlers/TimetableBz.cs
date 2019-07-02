using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports.ExcelHandlers {
    class TimetableBz : ExcelGeneral {
        public static string PerformData(DataTable dataTable) {
            Dictionary<string, Dictionary<string, ItemDoctor>> data = new Dictionary<string, Dictionary<string, ItemDoctor>>();
            ItemTimetable itemTimetable = new ItemTimetable();

            foreach (DataRow row in dataTable.Rows) {
                string filial = row["FILIAL"].ToString();
                string filname = row["FILNAME"].ToString();
                string dcode = row["DCODE"].ToString();
                string efio = row["EFIO"].ToString();
                string espec = row["ESPEC"].ToString();
                string dt = row["DT"].ToString();
                string time_start = row["TIME_START"].ToString();
                string time_end = row["TIME_END"].ToString();
                string free = row["FREE"].ToString();

                if (!itemTimetable.filials.ContainsKey(filial))
                    itemTimetable.filials.Add(filial, filname);

                if (!data.ContainsKey(filial))
                    data.Add(filial, new Dictionary<string, ItemDoctor>());

                if (!data[filial].ContainsKey(dcode))
                    data[filial].Add(dcode, new ItemDoctor(efio, espec));

                data[filial][dcode].cells.Add(new ItemCell(dt, time_start, time_end, free.Equals("0") ? false : true));
            }

            itemTimetable.filials.Add("data", data);

			string filePath = GetResultFilePath("bz_timetable.json", isPlainText:true);
			string fileContent = JsonConvert.SerializeObject(itemTimetable, Formatting.Indented);
			File.WriteAllText(filePath, fileContent);

            return filePath;
        }

        private class ItemTimetable {
            [JsonExtensionData]
            public Dictionary<string, object> filials = new Dictionary<string, object>();
        }

        private class ItemDoctor {
            public string efio;
            public string espec;
            public List<ItemCell> cells = new List<ItemCell>();

            public ItemDoctor(string efio, string espec) {
                this.efio = efio;
                this.espec = espec;
            }
        }

        private class ItemCell {
            public string dt;
            public string time_start;
            public string time_end;
            public bool free;
            public string room;

            public ItemCell(string dt, string time_start, string time_end, bool free, string room = "") {
                this.dt = dt;
                this.time_start = time_start;
                this.time_end = time_end;
                this.free = free;
                this.room = room;
            }
        }
    }
}
