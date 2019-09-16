using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class AverageCheck : ExcelGeneral {
		public static bool Process(string resultFile, Dictionary<string, object> periodCurrent, Dictionary<string, object> periodPrevious) {
			Logging.ToLog("Выполнение пост-обработки");

			string[] sheetNames = new string[] { "Факт", "Аванс", "Аванс_ЛМС" };
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws, sheetNames[0]))
				return false;

			string period1 = "период1";
			string period2 = "период2";

			try {
				period1 = GetPeriod(periodPrevious);
				period2 = GetPeriod(periodCurrent);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			foreach (string sheetName in sheetNames) {
				ws = wb.Sheets[sheetName];
				ws.Activate();

				try {
					CreateFormattingForSheet(xlApp, ws, sheetName, period1, period2);
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			}

			wb.Sheets[sheetNames[0]].Activate();
			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		#region Classes for data
		public class ItemAverageCheck {
			public Dictionary<string, ItemProgram> Programs { get; set; } = new Dictionary<string, ItemProgram>();
		}

		public class ItemProgram {
			public Dictionary<string, ItemFilial> Filials { get; set; } = new Dictionary<string, ItemFilial>();
			public Dictionary<string, ItemData> ProgramTotal { get; set; } = new Dictionary<string, ItemData>();
		}

		public class ItemFilial {
			public SortedDictionary<string, ItemGroup> Groups { get; set; } = new SortedDictionary<string, ItemGroup>();
			public Dictionary<string, ItemData> FilialTotals { get; set; } = new Dictionary<string, ItemData>();
		}

		public class ItemGroup {
			public SortedDictionary<string, ItemDepartment> Departmens { get; set; } = new SortedDictionary<string, ItemDepartment>();
			public Dictionary<string, ItemData> GroupTotals { get; set; } = new Dictionary<string, ItemData>();
		}

		public class ItemDepartment {
			public Dictionary<string, ItemData> Channels { get; set; } = new Dictionary<string, ItemData>();
		}

		public class ItemData {
			public double? CurrentCost { get; set; }
			public double? CurrentDiscountedCost { get; set; }
			public int? CurrentServicesCount { get; set; }
			public int? CurrentUniqPatientsCount { get; set; }
			public int? CurrentTreatmentsCount { get; set; }

			public double? PreviousCost { get; set; }
			public double? PreviousDiscountedCost { get; set; }
			public int? PreviousServicesCount { get; set; }
			public int? PreviousUniqPatientsCount { get; set; }
			public int? PreviousTreatmentsCount { get; set; }
		}
		#endregion

		public static ItemAverageCheck PerformData(DataTable dataTableCurrent, DataTable dataTablePrevious) {
			Logging.ToLog("Обработка данных");

			if (dataTableCurrent.Columns.Count != 11 &&
				dataTablePrevious.Columns.Count != 11) {
				Logging.ToLog("Невозможно выполнить обработку таблицы, кол-во столбцов не равно 11");
				return null;
			}

			ItemAverageCheck averageCheck = new ItemAverageCheck();

			ParseDataTable(dataTableCurrent, averageCheck, true);
			ParseDataTable(dataTablePrevious, averageCheck, false);

			return averageCheck;
		}

		private static void ParseDataTable(DataTable dataTable, ItemAverageCheck averageCheck, bool isCurrent) {
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

					if (channel.Equals("Физики факт")) {
						channel = "Физики";
						program = "Факт";
					} else if (channel.Equals("Физики аванс")) {
						channel = "Физики";
						program = "Аванс";
					}

					ItemData itemData;

					if (!averageCheck.Programs.ContainsKey(program))
						averageCheck.Programs.Add(program, new ItemProgram());

					if (!string.IsNullOrEmpty(filial) &&
						!averageCheck.Programs[program].Filials.ContainsKey(filial))
						averageCheck.Programs[program].Filials.Add(filial, new ItemFilial());

					if (!string.IsNullOrEmpty(group) &&
						!averageCheck.Programs[program].Filials[filial].Groups.ContainsKey(group))
						averageCheck.Programs[program].Filials[filial].Groups.Add(group, new ItemGroup());

					if (!string.IsNullOrEmpty(department) &&
						!averageCheck.Programs[program].Filials[filial].Groups[group].Departmens.ContainsKey(department))
						averageCheck.Programs[program].Filials[filial].Groups[group].Departmens.Add(department, new ItemDepartment());

					if (grpType == 0) {
						if (!averageCheck.Programs[program].Filials[filial].Groups[group].Departmens[department].Channels.ContainsKey(channel))
							averageCheck.Programs[program].Filials[filial].Groups[group].Departmens[department].Channels.Add(channel, new ItemData());

						itemData = averageCheck.Programs[program].Filials[filial].Groups[group].Departmens[department].Channels[channel];
					} else if (grpType == 1) {
						if (!averageCheck.Programs[program].Filials[filial].Groups[group].GroupTotals.ContainsKey(channel))
							averageCheck.Programs[program].Filials[filial].Groups[group].GroupTotals.Add(channel, new ItemData());

						itemData = averageCheck.Programs[program].Filials[filial].Groups[group].GroupTotals[channel];
					} else if (grpType == 2) {
						if (!averageCheck.Programs[program].Filials[filial].FilialTotals.ContainsKey(channel))
							averageCheck.Programs[program].Filials[filial].FilialTotals.Add(channel, new ItemData());

						itemData = averageCheck.Programs[program].Filials[filial].FilialTotals[channel];
					} else if (grpType == 3) {
						if (!averageCheck.Programs[program].ProgramTotal.ContainsKey(channel))
							averageCheck.Programs[program].ProgramTotal.Add(channel, new ItemData());

						itemData = averageCheck.Programs[program].ProgramTotal[channel];
					} else {
						Logging.ToLog("Неизвестный тип группировки - " + grpType);
						continue;
					}

					if (isCurrent) {
						itemData.CurrentServicesCount = servicesCount;
						itemData.CurrentDiscountedCost = discountedCost;
						itemData.CurrentCost = cost;
						itemData.CurrentUniqPatientsCount = uniqPatientsCount;
						itemData.CurrentTreatmentsCount = treatmentsCount;
					} else {
						itemData.PreviousServicesCount = servicesCount;
						itemData.PreviousDiscountedCost = discountedCost;
						itemData.PreviousCost = cost;
						itemData.PreviousUniqPatientsCount = uniqPatientsCount;
						itemData.PreviousTreatmentsCount = treatmentsCount;
					}
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}

				#region Старая чистка
				//try {
				//	string filial = dataRow["FILIAL"].ToString().ToUpper();
				//	if (filial.Contains("КУТУЗ")) {
				//		dataRow["FILIAL"] = "МДМ";

				//		if (dataTable.Columns.Contains("KYTYZOVSKIY"))
				//			dataRow["KYTYZOVSKIY"] = "Кутузовский";
				//	}

				//	string channelCleared = string.Empty;
				//	string programType = string.Empty;

				//	switch (dataRow["CTYPE"].ToString().ToUpper()) {
				//		case "ДРУГИЕ СК":
				//		case "ИНГОССТРАХ":
				//			channelCleared = "ДМС";
				//			break;
				//		case "ОМС":
				//			channelCleared = "ОМС";
				//			break;
				//		case "ЛМС 0":
				//			channelCleared = "ЛМС_0";
				//			programType = "Аванс";
				//			break;
				//		case "ЛМС 6":
				//			channelCleared = "ЛМС_6";
				//			programType = "Аванс";

				//			if (dataTable.Columns.Contains("DEP")) {
				//				string dep = dataRow["DEP"].ToString().ToUpper();
				//				if (dep.Equals("КОММЕРЧЕСКИЙ ОТДЕЛ") ||
				//					dep.Equals("РЕГИСТРАТУРА ВЗРОСЛАЯ") ||
				//					dep.Equals("РЕГИСТРАТУРА ДЕТСКАЯ"))
				//					programType = "Факт";
				//			}

				//			break;
				//		case "ФЛ ПРОГРАММЫ":
				//			channelCleared = "Физики";
				//			programType = "Аванс";
				//			break;
				//		case "ЧАСТНЫЕ КЛИЕНТЫ":
				//			channelCleared = "Физики";
				//			programType = "Факт";
				//			break;
				//		default:
				//			break;
				//	}

				//	if (dataTable.Columns.Contains("INDICATOR_BLOCK"))
				//		if (dataRow["INDICATOR_BLOCK"].ToString().ToUpper().Equals("КОММЕРЧЕСКИЙ ОТДЕЛ")) {
				//			channelCleared = "Физики";
				//			programType = "Аванс";
				//		}

				//	dataRow["CHA_CLEARED"] = channelCleared;

				//	if (!string.IsNullOrEmpty(programType))
				//		dataRow["ATYPE"] = programType;

				//	if (filial.Equals("К-УРАЛ")) {
				//		if (dataTable.Columns.Contains("DEP"))
				//			switch (dataRow["DEP"].ToString().ToUpper()) {
				//				case "КОММЕРЧЕСКИЙ ОТДЕЛ":
				//				case "ПРОФПАТОЛОГ":
				//					dataRow["DEP"] = "ПРОФОСМОТР";
				//					break;
				//				default:
				//					break;
				//			}

				//		//if (dataRow["ATYPE"].ToString().ToUpper().Equals("АВАНС"))
				//		//	dataRow["ATYPE"] = "АвансКУР";
				//	}
				//} catch (Exception e) {
				//	Console.WriteLine(e.Message);
				//}
				#endregion
			}
		}



		public static string WriteAverageCheckToExcel(ItemAverageCheck averageCheck, string resultFilePrefix, string templateFileName) {
			IWorkbook workbook = null;
			ISheet sheet = null;
			string resultFile = string.Empty;

			Logging.ToLog("Запись данных в книгу Excel");

			string[] sheetNames = new string[] { "Факт", "Аванс", "Аванс_ЛМС" };
			if (!CreateNewIWorkbook(resultFilePrefix, templateFileName,
					out workbook, out sheet, out resultFile, sheetNames[0]))
				return string.Empty;

			foreach (string sheetName in sheetNames) {
				sheet = workbook.GetSheet(sheetName);

				string program = string.Empty;
				string channelLeft = string.Empty;
				string channelRight = string.Empty;

				if (sheetName.Equals("Факт")) {
					program = "Факт";
					channelLeft = "ДМС";
					channelRight = "Физики";
				} else if (sheetName.Equals("Аванс")) {
					program = "Аванс";
					channelLeft = "ДМС";
					channelRight = "ЛМС 6";
				} else if (sheetName.Equals("Аванс_ЛМС")) {
					program = "Аванс";
					channelLeft = "ЛМС 0";
				} else {
					Logging.ToLog("Неизвестное имя листа: " + sheetName);
					continue;
				}

				if (!averageCheck.Programs.ContainsKey(program)) {
					Logging.ToLog("Блок данных не содержит программу: " + program);
					continue;
				}

				int rowNumber = 6;

				foreach (KeyValuePair<string, ItemFilial> filial in averageCheck.Programs[program].Filials) {
					foreach (KeyValuePair<string, ItemGroup> group in filial.Value.Groups) {
						foreach (KeyValuePair<string, ItemDepartment> department in group.Value.Departmens) {
							ItemData dataLeft = null;
							if (department.Value.Channels.ContainsKey(channelLeft))
								dataLeft = department.Value.Channels[channelLeft];

							ItemData dataRight = null;
							if (department.Value.Channels.ContainsKey(channelRight))
								dataRight = department.Value.Channels[channelRight];

							object[] values;

							if (dataLeft != null || dataRight != null)
								try {
									values = GenerateValuesToWrite(sheetName,
										   filial.Key,
										   group.Key,
										   department.Key,
										   dataLeft ?? new ItemData(),
										   dataRight ?? new ItemData());
									WriteOutValues(values, sheet, ref rowNumber);
								} catch (Exception e) {
									Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
								}
						}

						ItemData groupTotalLeft = null;
						if (group.Value.GroupTotals.ContainsKey(channelLeft))
							groupTotalLeft = group.Value.GroupTotals[channelLeft];

						ItemData groupTotalRight = null;
						if (group.Value.GroupTotals.ContainsKey(channelRight))
							groupTotalRight = group.Value.GroupTotals[channelRight];

						object[] groupTotalValues;

						if (groupTotalLeft != null || groupTotalRight != null)
							try {
								groupTotalValues = GenerateValuesToWrite(sheetName,
													filial.Key,
													group.Key + " - ИТОГО",
													string.Empty,
													groupTotalLeft ?? new ItemData(),
													groupTotalRight ?? new ItemData());
								WriteOutValues(groupTotalValues, sheet, ref rowNumber);
							} catch (Exception e) {
								Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
							}
					}

					ItemData filialTotalLeft = null;
					if (filial.Value.FilialTotals.ContainsKey(channelLeft))
						filialTotalLeft = filial.Value.FilialTotals[channelLeft];

					ItemData filialTotalRight = null;
					if (filial.Value.FilialTotals.ContainsKey(channelRight))
						filialTotalRight = filial.Value.FilialTotals[channelRight];

					object[] filialTotalValues;

					if (filialTotalLeft != null || filialTotalRight != null)
						try {
							filialTotalValues = GenerateValuesToWrite(sheetName,
												filial.Key + " - ИТОГО",
												string.Empty,
												string.Empty,
												filialTotalLeft ?? new ItemData(),
												filialTotalRight ?? new ItemData());
							WriteOutValues(filialTotalValues, sheet, ref rowNumber);
						} catch (Exception e) {
							Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
						}
				}

				ItemData programTotalLeft = null;
				if (averageCheck.Programs[program].ProgramTotal.ContainsKey(channelLeft))
					programTotalLeft = averageCheck.Programs[program].ProgramTotal[channelLeft];

				ItemData programTotalRight = null;
				if (averageCheck.Programs[program].ProgramTotal.ContainsKey(channelRight))
					programTotalRight = averageCheck.Programs[program].ProgramTotal[channelRight];

				object[] programTotalValues;

				if (programTotalLeft != null || programTotalRight != null)
					try {
						programTotalValues = GenerateValuesToWrite(sheetName,
											"ИТОГО",
											string.Empty,
											string.Empty,
											programTotalLeft ?? new ItemData(),
											programTotalRight ?? new ItemData());
						WriteOutValues(programTotalValues, sheet, ref rowNumber);
					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					}
			}

			if (!SaveAndCloseIWorkbook(workbook, resultFile))
				return string.Empty;

			return resultFile;
		}

		private static object[] GenerateValuesToWrite(string sheetName,
												string filial,
												string group,
												string department,
												ItemData dataLeft,
												ItemData dataRight) {
			object[] values;

			if (sheetName.Equals("Факт") ||
				sheetName.Equals("Аванс")) {
				values = new object[] {
					filial,
					group,
					department,
					dataLeft.PreviousDiscountedCost ?? null,
					dataRight.PreviousDiscountedCost ?? null,
					dataLeft.CurrentDiscountedCost ?? null,
					dataRight.CurrentDiscountedCost ?? null,
					null,
					null,
					dataLeft.PreviousServicesCount ?? null,
					dataRight.PreviousServicesCount ?? null,
					dataLeft.CurrentServicesCount ?? null,
					dataRight.CurrentServicesCount ?? null,
					dataLeft.PreviousUniqPatientsCount ?? null,
					dataRight.PreviousUniqPatientsCount ?? null,
					dataLeft.CurrentUniqPatientsCount ?? null,
					dataRight.CurrentUniqPatientsCount ?? null
				};
			} else if (sheetName.Equals("Аванс_ЛМС")) {
				values = new object[] {
					filial,
					group,
					department,
					dataLeft.PreviousDiscountedCost ?? null,
					dataLeft.CurrentDiscountedCost ?? null,
					null,
					dataLeft.PreviousServicesCount ?? null,
					dataLeft.CurrentServicesCount ?? null,
					dataLeft.PreviousUniqPatientsCount ?? null,
					dataLeft.CurrentUniqPatientsCount ?? null,
				};
			} else
				throw new Exception("Неизвестное имя листа: " + sheetName);

			return values;
		}

		public static void WriteOutValues(object[] values, ISheet sheet, ref int rowNumber) {
			int columnNumber = 0;

			IRow row = null;
			try { row = sheet.GetRow(rowNumber); } catch (Exception) { }

			if (row == null)
				row = sheet.CreateRow(rowNumber);

			foreach (object value in values) {
				ICell cell = null;
				try { cell = row.GetCell(columnNumber); } catch (Exception) { }

				if (cell == null)
					cell = row.CreateCell(columnNumber);

				if (value is string)
					cell.SetCellValue((string)value);
				else if (value is double)
					cell.SetCellValue((double)value);
				else if (value is int)
					cell.SetCellValue((int)value);

				columnNumber++;
			}

			rowNumber++;
		}

		public static string GetPeriod(Dictionary<string, object> period) {
			return period["@dateBegin"].ToString() + "-" + period["@dateEnd"].ToString();
		}

		private static void CreateFormattingForSheet(Excel.Application xlApp, Excel.Worksheet ws, string sheetName, string period1, string period2) {
			Logging.ToLog("Подготовка листа: " + sheetName);

			ws.Activate();

			ws.Cells.Replace("date1", period1, Excel.XlLookAt.xlWhole);
			ws.Cells.Replace("date2", period2, Excel.XlLookAt.xlWhole);

			int usedRows = ws.UsedRange.Rows.Count;
			int usedColumns = ws.UsedRange.Columns.Count;
			string lastColumn = ColumnIndexToColumnLetter(usedColumns);

			ws.Range["A7: " + lastColumn + "7"].Select();
			xlApp.Selection.AutoFill(ws.Range["A7:" + lastColumn + usedRows], Excel.XlAutoFillType.xlFillFormats);
			AddBoldBorder(ws.Range["A7:" + lastColumn + usedRows]);

			string[] rangesWithFormula = new string[] { "H7:I7", "R7:Y7" };
			string columnsToHide = "J:M";
			int scrollColumn = 4;
			if (sheetName.Equals("Аванс_ЛМС")) {
				rangesWithFormula = new string[] { "F7:F7", "K7:N7" };
				columnsToHide = "G:H";
				scrollColumn = 2;
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

			ws.Columns[columnsToHide].Select();
			xlApp.Selection.EntireColumn.Hidden = true;
			xlApp.ActiveWindow.ScrollColumn = scrollColumn;

			for (int row = 7; row <= usedRows; row++) {
				string filial = ws.Range["A" + row].Value2;
				string block = ws.Range["B" + row].Value2;

				double tintAndShade = -1;
				if (!string.IsNullOrEmpty(filial) && filial.Contains("ИТОГ"))
					tintAndShade = 0.599993896298105;
				else if (!string.IsNullOrEmpty(block) && block.Contains("ИТОГ"))
					tintAndShade = 0.799981688894314;

				if (tintAndShade != -1) {
					Excel.Range rangeToColorize = ws.Range["A" + row + ":" + lastColumn + row];
					rangeToColorize.Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
					rangeToColorize.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4;
					rangeToColorize.Interior.TintAndShade = tintAndShade;
					rangeToColorize.Interior.PatternTintAndShade = 0;
				}
			}

			ws.Range["A1"].Select();
		}
	}
}
