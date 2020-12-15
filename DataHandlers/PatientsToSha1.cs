using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace MISReports.DataHandlers {
	public class PatientsToSha1 {
		public static DataTable PerformDataTable(DataTable dataTable) {
			//DataRow dataRowIvanov = dataTable.NewRow();
			//dataRowIvanov["PCODE"] = "0";
			//dataRowIvanov["FULLNAME"] = "Иванов Иван Иванович";
			//dataRowIvanov["BDATE"] = "01.01.1980";
			//dataRowIvanov["PASPSER"] = "56 31";
			//dataRowIvanov["PASPNUM"] = "123456";
			//dataRowIvanov["MAX"] = "01.01.2010";
			//dataTable.Rows.InsertAt(dataRowIvanov, 0);


			//select
			//c.pcode
			//, c.fullname
			//, c.bdate
			//, c.PASPSER
			//, c.PASPNUM
			//, max(cl.fdate)
			//from clients c
			//left
			//join clhistnum cl on cl.pcode = c.pcode
			//where extract(year from c.bdate) >= 2000
			//group by 1,2,3,4,5



			DataTable dataTableResult = new DataTable();
			dataTableResult.Clear();
			dataTableResult.Columns.Add("SOURCE");
			dataTableResult.Columns.Add("CLIENTID");
			dataTableResult.Columns.Add("HASH1");
			//dataTableResult.Columns.Add("HASH2");
			dataTableResult.Columns.Add("STATE");

			dataTableResult.Columns.Add("product1");
			dataTableResult.Columns.Add("product2");
			dataTableResult.Columns.Add("product3");
			dataTableResult.Columns.Add("product4");
			dataTableResult.Columns.Add("product5");
			dataTableResult.Columns.Add("product6");
			dataTableResult.Columns.Add("product7");
			dataTableResult.Columns.Add("product8");
			dataTableResult.Columns.Add("product9");
			dataTableResult.Columns.Add("product10");
			dataTableResult.Columns.Add("product11");
			dataTableResult.Columns.Add("product12");
			dataTableResult.Columns.Add("product13");
			dataTableResult.Columns.Add("product14");
			dataTableResult.Columns.Add("product15");
			dataTableResult.Columns.Add("product16");
			dataTableResult.Columns.Add("product17");
			dataTableResult.Columns.Add("product18");
			dataTableResult.Columns.Add("product_other");

			string source = "BZ";

			foreach (DataRow row in dataTable.Rows) {
				//string minSectid = row["MIN_SECTID"].ToString();
				string maxFdateHistnim = row["MAX_FDATE_HISTNUM"].ToString();

				if (!string.IsNullOrEmpty(maxFdateHistnim)) {
					//if (!minSectid.Equals("4363"))
					//	continue;
				//} else {
					if (DateTime.TryParse(maxFdateHistnim, out DateTime resultParse)) {
						if (resultParse < DateTime.Now)// && !minSectid.Equals("4363"))
							continue;
					}
				}

				string pcode = row["PCODE"].ToString();
				string fullname = row["FULLNAME"].ToString();
				string bdate = row["BDATE"].ToString().Replace(" 0:00:00", "");
				string paspser = row["PASPSER"].ToString();
				string paspnum = row["PASPNUM"].ToString();
				//string max = row["MAX"].ToString().Replace(" 0:00:00", "");

				string hash1Original = fullname + bdate + paspser + paspnum;
				//string hash2Original = fullname + bdate;

				string hash1 = GetSha1Hash(GetNormalizedString(hash1Original));
				//string hash2 = GetSha1Hash(GetNormalizedString(hash2Original));

				string departments = row["LIST"].ToString();
				if (string.IsNullOrEmpty(departments))
					continue;


				int state = 2;
				//if (!string.IsNullOrEmpty(max)) {
				//	DateTime maxDate = DateTime.Parse(max);
				//	if (maxDate.Date >= DateTime.Now)
				//		state = 2;
				//	else if ((DateTime.Now - maxDate).TotalDays / 365 <= 5)
				//		state = 1;
				//}

				DataRow dataRow = dataTableResult.NewRow();
				dataRow["SOURCE"] = source;
				dataRow["CLIENTID"] = pcode;
				dataRow["HASH1"] = hash1 + GetNormalizedString(hash1Original).Substring(0, 3).ToUpper();
				//dataRow["HASH2"] = hash2 + GetNormalizedString(hash2Original).Substring(1, 3).ToUpper();
				dataRow["STATE"] = state;
				dataTableResult.Rows.Add(dataRow);

				Dictionary<string, int> departs = new Dictionary<string, int>();

				string[] departmentsSplitted = departments.Split(',');
				foreach (string dept in departmentsSplitted) {
					string[] splitted = dept.Split(';');

					if (!departs.ContainsKey(splitted[0]))
						departs.Add(splitted[0], int.Parse(splitted[1]));
				}

				Dictionary<string, string> products = new Dictionary<string, string> {
					{ "949", "1" }, //СТОМАТОЛОГИЯ
					{ "772", "2" }, //ЛАБОРАТОРИЯ
					{ "770", "3" }, //УЗИ
					{ "820", "4" }, //РЕНТГЕН
					{ "746", "5" }, //ГИНЕКОЛОГИЯ
					{ "732", "6" }, //ТЕРАПИЯ
					{ "741", "7" }, //ОТОРИНОЛАРИНГОЛОГИЯ
					{ "742", "8" }, //НЕВРОЛОГИЯ
					{ "991309905", "9" }, //ФИЗИОПРОЦЕДУРЫ
					{ "740", "10" }, //ОФТАЛЬМОЛОГИЯ
					{ "991330843", "11" }, //СКОРАЯ МЕДИЦИНСКАЯ ПОМОЩЬ
					{ "765", "12" }, //ТРАВМАТОЛОГИЯ
					{ "20001103", "13" }, //МАССАЖ
					{ "738", "14" }, //ХИРУРГИЯ
					{ "821", "15" }, //ПРОЦЕДУРНЫЙ КАБИНЕТ
					{ "991328713", "16" }, //КАРДИОЛОГИЯ
					{ "739", "17" }, //УРОЛОГИЯ
					{ "990319830", "18" }, //ГАСТРОЭНТЕРОЛОГИЯ
					{ "прочие", "_other" } //ПРОЧИЕ
				};

				foreach (KeyValuePair<string, string> prod in products) {

					if (departs.ContainsKey(prod.Key)) {
						dataRow["product" + prod.Value] = departs[prod.Key];
						departs[prod.Key] = 0;
					}

					if (prod.Key.Equals("прочие"))
						foreach (int item in departs.Values) {
							int currentValue = 0;
							string curVal = dataRow["product" + prod.Value].ToString();
							if (!string.IsNullOrEmpty(curVal))
								int.TryParse(curVal, out currentValue);

							dataRow["product" + prod.Value] = currentValue + item;
						}
				}

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
