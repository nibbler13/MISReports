using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class CompetitiveGroups : ExcelGeneral {
		private static readonly string[] sheetNames = new string[] {
			"Факт МСК",
			"Аванс МСК",
			"Факт СПБ+Уфа",
			"Аванс СПБ+Уфа",
			"Факт КРД+Сочи+КУрал+Казань",
			"Аванс КРД+Сочи+КУрал+Казань"
		};

		public static bool Process(string resultFile, Dictionary<string, object> parameters) {
			Logging.ToLog("Выполнение пост-обработки");

			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws, sheetNames[0]))
				return false;

			string period = "период";

			try {
				period = AverageCheck.GetPeriod(parameters);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			foreach (string sheetName in sheetNames) {
				ws = wb.Sheets[sheetName];
				ws.Activate();

				try {
					CreateFormattingForSheet(xlApp, ws, sheetName, period);
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			}

			wb.Sheets[sheetNames[0]].Activate();
			SaveAndCloseWorkbook(xlApp, wb, ws);
			return true;
		}

		private static void CreateFormattingForSheet(Excel.Application xlApp, Excel.Worksheet ws, string sheetName, string period) {
			Logging.ToLog("Подготовка листа: " + sheetName);

			ws.Activate();

			ws.Cells.Replace("date", period, Excel.XlLookAt.xlWhole);

			int usedRows = ws.UsedRange.Rows.Count;
			int usedColumns = ws.UsedRange.Columns.Count;
			string lastColumn = ColumnIndexToColumnLetter(usedColumns);

			ws.Range["A7: " + lastColumn + "7"].Select();
			xlApp.Selection.AutoFill(ws.Range["A7:" + lastColumn + usedRows], Excel.XlAutoFillType.xlFillFormats);
			AddBoldBorder(ws.Range["A7:" + lastColumn + usedRows]);

			string[] rangesWithFormula;
			string columnsToHide;

			if (sheetName.Contains("МСК")) {
				rangesWithFormula = new string[] { "AQ7:BJ7" };
				columnsToHide = "M:V";

			} else if (sheetName.Contains("СПБ+Уфа")) {
				rangesWithFormula = new string[] { "S7:Z7" };
				columnsToHide = "G:J";

			} else if (sheetName.Contains("КРД+Сочи+КУрал+Казань")) {
				rangesWithFormula = new string[] { "AI7:AX7" };
				columnsToHide = "K:R";

			} else {
				Logging.ToLog("Неизвестное имя листа: " + sheetName);
				return;
			}

			foreach (string rangeWithFormula in rangesWithFormula) {
				ws.Range[rangeWithFormula].Select();
				string rangeToFill = rangeWithFormula.Substring(0, rangeWithFormula.Length - 1) + usedRows;
				xlApp.Selection.AutoFill(
					ws.Range[rangeToFill],
					Excel.XlAutoFillType.xlFillValues);

				string rangeToFillFrom = rangeWithFormula.Replace("7", "8");
				rangeToFill = rangeWithFormula.Substring(0, rangeWithFormula.Length - 1) + "8";
				ws.Range[rangeToFillFrom].Select();
				xlApp.Selection.AutoFill(
					ws.Range[rangeToFill],
					Excel.XlAutoFillType.xlFillValues);
			}

			for (int row = 7; row <= usedRows; row++) {
				string block = ws.Range["A" + row].Value2;

				double tintAndShade = -1;
				if (!string.IsNullOrEmpty(block) && block.Contains("ИТОГ"))
					tintAndShade = 0.799981688894314;
				
				if (!string.IsNullOrEmpty(block) && block.Equals("ИТОГО"))
					tintAndShade = 0.599993896298105;

				if (tintAndShade != -1) {
					Excel.Range rangeToColorize = ws.Range["A" + row + ":" + lastColumn + row];
					rangeToColorize.Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
					rangeToColorize.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4;
					rangeToColorize.Interior.TintAndShade = tintAndShade;
					rangeToColorize.Interior.PatternTintAndShade = 0;
				}
			}

			ws.Columns["C:" + lastColumn].Select();
			xlApp.Selection.Columns.AutoFit();

			ws.Columns[columnsToHide].Select();
			xlApp.Selection.EntireColumn.Hidden = true;

			ws.Range["A1"].Select();
		}


		#region Classes for data
		public class ItemCompetitiveGroups {
			public Dictionary<string, ItemProgram> Programs { get; set; } = new Dictionary<string, ItemProgram>();
		}

		public class ItemProgram {
			public Dictionary<string, ItemGroup> Groups { get; set; } = new Dictionary<string, ItemGroup>();
			public Dictionary<string, ItemFilial> ProgramTotals { get; set; } = new Dictionary<string, ItemFilial>();
		}

		public class ItemGroup {
			public SortedDictionary<string, ItemDepartment> Departments { get; set; } = new SortedDictionary<string, ItemDepartment>();
			public Dictionary<string, ItemFilial> GroupTotals { get; set; } = new Dictionary<string, ItemFilial>();
		}

		public class ItemDepartment {
			public Dictionary<string, ItemFilial> Filials { get; set; } = new Dictionary<string, ItemFilial>();
		}

		public class ItemFilial {
			public SortedDictionary<string, ItemData> Channels { get; set; } = new SortedDictionary<string, ItemData>();
		}

		public class ItemData {
			public double? Cost { get; set; }
			public double? DiscountedCost { get; set; }
			public int? ServicesCount { get; set; }
			public int? UniqPatientsFirstTime { get; set; }
			public int? UniqPatientsCount { get; set; }
			public int? TreatmentsCount { get; set; }
		}
		#endregion

		public static ItemCompetitiveGroups PerformData(DataTable dataTable) {
			Logging.ToLog("Обработка данных");

			if (dataTable.Columns.Count != 11) {
				Logging.ToLog("Невозможно выполнить обработку таблицы, кол-во столбцов не равно 11");
				return null;
			}

			ItemCompetitiveGroups competitiveGroups = new ItemCompetitiveGroups();
			ParseDataTable(dataTable, competitiveGroups);

			return competitiveGroups;
		}

		private static void ParseDataTable(DataTable dataTable, ItemCompetitiveGroups competitiveGroups) {
			foreach (DataRow dataRow in dataTable.Rows) {
				try {
					int grpType = Convert.ToInt32(dataRow["GRPTYPE"].ToString());
					string filial = dataRow["FILIAL"].ToString();
					string group = dataRow["LONGTEXT1"].ToString();
					string department = dataRow["DEPART"].ToString();
					string channel = dataRow["CHANEL_TYPE"].ToString();
					string program = dataRow["PRG_TYPE"].ToString();
					double cost = Convert.ToDouble(dataRow["SUM_SERV"].ToString());
					int servicesCount = Convert.ToInt32(dataRow["COUNT_SERV"].ToString());
					int uniqPatientsCount = Convert.ToInt32(dataRow["UNI_PAC"].ToString());
					int treatmentsCount = Convert.ToInt32(dataRow["UNI_TREAT"].ToString());
					double discountedCost = Convert.ToDouble(dataRow["DISC_SUM_SERV"].ToString());

					if (filial.Equals("СУЩ") && department.Equals("ГАСТРОЭНТЕРОЛОГИЯ"))
						Console.WriteLine("");

					bool isFirstVisit = false;
					if (channel.Contains("перв"))
						isFirstVisit = true;

					if (channel.StartsWith("Физики факт")) {
						channel = "Физики";
						program = "Факт";
					} else if (channel.StartsWith("Физики аванс")) {
						channel = "Физики";
						program = "Аванс";
					}

					ItemData itemData;

					if (!competitiveGroups.Programs.ContainsKey(program))
						competitiveGroups.Programs.Add(program, new ItemProgram());

					if (!string.IsNullOrEmpty(group) &&
						!competitiveGroups.Programs[program].Groups.ContainsKey(group))
						competitiveGroups.Programs[program].Groups.Add(group, new ItemGroup());

					if (!string.IsNullOrEmpty(department) &&
						!competitiveGroups.Programs[program].Groups[group].Departments.ContainsKey(department))
						competitiveGroups.Programs[program].Groups[group].Departments.Add(department, new ItemDepartment());

					if (!string.IsNullOrEmpty(filial) && !string.IsNullOrEmpty(department) &&
						!competitiveGroups.Programs[program].Groups[group].Departments[department].Filials.ContainsKey(filial))
						competitiveGroups.Programs[program].Groups[group].Departments[department].Filials.Add(filial, new ItemFilial());

					if (grpType == 0) {
						if (!competitiveGroups.Programs[program].Groups[group].Departments[department].Filials[filial].Channels.ContainsKey(channel))
							competitiveGroups.Programs[program].Groups[group].Departments[department].Filials[filial].Channels.Add(channel, new ItemData());

						itemData = competitiveGroups.Programs[program].Groups[group].Departments[department].Filials[filial].Channels[channel];

					} else if (grpType == 1) {
						if (!competitiveGroups.Programs[program].Groups[group].GroupTotals.ContainsKey(filial))
							competitiveGroups.Programs[program].Groups[group].GroupTotals.Add(filial, new ItemFilial());

						if (!competitiveGroups.Programs[program].Groups[group].GroupTotals[filial].Channels.ContainsKey(channel))
							competitiveGroups.Programs[program].Groups[group].GroupTotals[filial].Channels.Add(channel, new ItemData());

						itemData = competitiveGroups.Programs[program].Groups[group].GroupTotals[filial].Channels[channel];

					} else if (grpType == 2) {
						if (!competitiveGroups.Programs[program].ProgramTotals.ContainsKey(filial))
							competitiveGroups.Programs[program].ProgramTotals.Add(filial, new ItemFilial());

						if (!competitiveGroups.Programs[program].ProgramTotals[filial].Channels.ContainsKey(channel))
							competitiveGroups.Programs[program].ProgramTotals[filial].Channels.Add(channel, new ItemData());

						itemData = competitiveGroups.Programs[program].ProgramTotals[filial].Channels[channel];

					} else if (grpType == 3) {
						continue;

					} else {
						Logging.ToLog("Неизвестный тип группировки - " + grpType);
						continue;
					}

					if (itemData.ServicesCount.HasValue)
						itemData.ServicesCount += servicesCount;
					else
						itemData.ServicesCount = servicesCount;

					if (itemData.DiscountedCost.HasValue)
						itemData.DiscountedCost += discountedCost;
					else
						itemData.DiscountedCost = discountedCost;

					if (itemData.Cost.HasValue)
						itemData.Cost += cost;
					else
						itemData.Cost = cost;

					if (isFirstVisit)
						itemData.UniqPatientsFirstTime = uniqPatientsCount;
					else
						itemData.UniqPatientsCount = uniqPatientsCount;

					if (itemData.TreatmentsCount.HasValue)
						itemData.TreatmentsCount += treatmentsCount;
					else
						itemData.TreatmentsCount = treatmentsCount;
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			}
		}


		public static string WriteAverageCheckToExcel(ItemCompetitiveGroups competitiveGroups, string resultFilePrefix, string templateFileName) {
			IWorkbook workbook = null;
			ISheet sheet = null;
			string resultFile = string.Empty;

			Logging.ToLog("Запись данных в книгу Excel");

			if (!CreateNewIWorkbook(resultFilePrefix, templateFileName,
					out workbook, out sheet, out resultFile, sheetNames[0]))
				return string.Empty;

			foreach (string sheetName in sheetNames) {
				sheet = workbook.GetSheet(sheetName);

				string program = string.Empty;
				string[] filials = null;
				string channelLeft = string.Empty;
				string channelRight = string.Empty;

				if (sheetName.Contains("Факт")) {
					program = "Факт";
					channelLeft = "ДМС";
					channelRight = "Физики";
				} else if (sheetName.Contains("Аванс")) {
					program = "Аванс";
					channelLeft = "ДМС";
					channelRight = "ЛМС 6";
				} else {
					Logging.ToLog("Неизвестное имя листа: " + sheetName);
					continue;
				}

				if (sheetName.Contains("МСК")) 
					filials = new string[] { "М-СРЕТ", "МДМ", "СУЩ", "КУТУЗ", "СКОРАЯ" };

				else if (sheetName.Contains("СПБ+Уфа"))
					filials = new string[] { "С-Пб.", "Уфа" };

				else if (sheetName.Contains("КРД+Сочи+КУрал+Казань"))
					filials = new string[] { "Красн", "Сочи", "К-УРАЛ", "Казань" };

				else {
					Logging.ToLog("Неизвестное имя листа: " + sheetName);
					continue;
				}

				if (!competitiveGroups.Programs.ContainsKey(program)) {
					Logging.ToLog("Блок данных не содержит программу: " + program);
					continue;
				}

				int rowNumber = 6;

				foreach (KeyValuePair<string, ItemGroup> groupNew in competitiveGroups.Programs[program].Groups) {
					foreach (KeyValuePair<string, ItemDepartment> departmentNew in groupNew.Value.Departments)
						CheckAndWriteDataToRow(
							sheet,
							groupNew.Key,
							departmentNew.Key,
							filials, 
							channelLeft, 
							channelRight, 
							departmentNew.Value.Filials, 
							ref rowNumber);

					CheckAndWriteDataToRow(
						sheet,
						groupNew.Key + " - ИТОГО",
						string.Empty,
						filials,
						channelLeft,
						channelRight,
						competitiveGroups.Programs[program].Groups[groupNew.Key].GroupTotals,
						ref rowNumber);
				}

				CheckAndWriteDataToRow(
					sheet,
					"ИТОГО",
					string.Empty,
					filials,
					channelLeft,
					channelRight,
					competitiveGroups.Programs[program].ProgramTotals,
					ref rowNumber);
			}

			if (!SaveAndCloseIWorkbook(workbook, resultFile))
				return string.Empty;

			return resultFile;
		}

		private static void CheckAndWriteDataToRow(ISheet sheet,
							   string groupName,
							   string departmentName,
							   string[] filialNames,
							   string channelLeft,
							   string channelRight,
							   Dictionary<string, ItemFilial> filials,
							   ref int rowNumber) {
			List<Tuple<string, ItemData>> items = new List<Tuple<string, ItemData>>();
			bool hasData = false;

			foreach (string filialName in filialNames) {
				if (!filials.ContainsKey(filialName)) {
					items.Add(new Tuple<string, ItemData>(channelLeft, null));
					items.Add(new Tuple<string, ItemData>(channelRight, null));
					continue;
				}

				ItemFilial itemFilial = filials[filialName];

				if (itemFilial.Channels.ContainsKey(channelLeft)) {
					items.Add(new Tuple<string, ItemData> ( channelLeft, itemFilial.Channels[channelLeft] ));
					hasData = true;
				} else
					items.Add(new Tuple<string, ItemData>(channelLeft, null));

				if (itemFilial.Channels.ContainsKey(channelRight)) {
					items.Add(new Tuple<string, ItemData>(channelRight, itemFilial.Channels[channelRight]));
					hasData = true;
				} else
					items.Add(new Tuple<string, ItemData>(channelRight, null));
			}

			if (!hasData)
				return;

			try {
				object[] values = GenerateValuesToWrite(groupName, departmentName, items);
				AverageCheck.WriteOutValues(values, sheet, ref rowNumber);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}
		}



		private static object[] GenerateValuesToWrite(string group, string department, List<Tuple<string, ItemData>> items) {
			List<object> valuesDiscountedCost = new List<object>();
			List<object> valuesServicesCount = new List<object>();
			List<object> valuesUniquePatients = new List<object>();

			foreach (Tuple<string, ItemData> item in items) {
				if (item.Item2 == null) {
					valuesDiscountedCost.Add(null);
					valuesServicesCount.Add(null);
					valuesUniquePatients.Add(null);

					if (!item.Item1.Equals("ДМС")) {
						valuesUniquePatients.Add(null);
						valuesUniquePatients.Add(null);
					}

					continue;
				}

				valuesDiscountedCost.Add(item.Item2.DiscountedCost ?? null); //DiscountedCost
				valuesServicesCount.Add(item.Item2.ServicesCount ?? null);
				valuesUniquePatients.Add(item.Item2.UniqPatientsCount ?? null);

				if (!item.Item1.Equals("ДМС")) {
					valuesUniquePatients.Add(item.Item2.UniqPatientsFirstTime ?? null);

					if (item.Item2.UniqPatientsCount == null && item.Item2.UniqPatientsFirstTime == null)
						valuesUniquePatients.Add(null);
					else {
						int patientCount = item.Item2.UniqPatientsCount ?? 0;
						int patientFirstTime = item.Item2.UniqPatientsFirstTime ?? 0;
						valuesUniquePatients.Add(patientCount + patientFirstTime);
					}
				}
			}

			List<object> objects = new List<object>() { group, department };
			List<object>[] values = new List<object>[] { valuesDiscountedCost, valuesServicesCount, valuesUniquePatients };
			foreach (List<object> list in values)
				foreach (object value in list)
					objects.Add(value);

			return objects.ToArray();
		}
	}
}
