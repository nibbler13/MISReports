using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class RegistryMarks : ExcelGeneral {

		//============================ RegistryMarks ============================
		public static bool Process(
			string resultFile, DataTable dataTable, DateTime dateTimeBegin) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws))
				return false;

			try {
				ws.Columns["C:C"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ Ч:мм;@";
				ws.Range["A1"].Select();
				xlApp.Selection.AutoFilter();
				ws.UsedRange.AutoFilter(3, ">" + dateTimeBegin.ToOADate(), Excel.XlAutoFilterOperator.xlAnd);
				ws.UsedRange.AutoFilter(4, "Плохо");
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				RegistryMarksAddPivotTables(wb, xlApp, dataTable, dateTimeBegin);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			wb.Sheets[1].Name = "Негативные отзывы";
			wb.Sheets["Сводная таблица"].Activate();
			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static void RegistryMarksAddPivotTables(
			Excel.Workbook wb, Excel.Application xlApp, DataTable dataTable, DateTime dateTimeBegin) {
			SortedDictionary<string, ItemRegistryMark> marksSelectedPeriodByFilials =
				new SortedDictionary<string, ItemRegistryMark>();
			SortedDictionary<string, SortedDictionary<string, ItemRegistryMark>> marksByWeeks
				= new SortedDictionary<string, SortedDictionary<string, ItemRegistryMark>>();

			List<string> uniqueInnerKeys = new List<string>();

			foreach (DataRow dataRow in dataTable.Rows) {
				try {
					DateTime createDate = DateTime.Parse(dataRow["createdate"].ToString());

					int weekNumber = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(
						createDate, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday);
					string dictByWeeksInnerKey = createDate.Year + " / " + weekNumber;

					if (!uniqueInnerKeys.Contains(dictByWeeksInnerKey))
						uniqueInnerKeys.Add(dictByWeeksInnerKey);

					string shortname = dataRow["SHORTNAME"].ToString();
					string department = dataRow["DEPARTMENT"].ToString();
					string dictsOuterKey = shortname + " / " + department;

					string mark = dataRow["MARK"].ToString();

					if (!marksByWeeks.Keys.Contains(dictsOuterKey))
						marksByWeeks.Add(dictsOuterKey, new SortedDictionary<string, ItemRegistryMark>());

					if (!marksByWeeks[dictsOuterKey].Keys.Contains(dictByWeeksInnerKey))
						marksByWeeks[dictsOuterKey].Add(dictByWeeksInnerKey, new ItemRegistryMark(shortname, department));

					if (!marksSelectedPeriodByFilials.ContainsKey(dictsOuterKey))
						marksSelectedPeriodByFilials.Add(dictsOuterKey, new ItemRegistryMark(shortname, department));

					if (mark.Contains("Плохо")) {
						marksByWeeks[dictsOuterKey][dictByWeeksInnerKey].MarkBad++;

						if (createDate >= dateTimeBegin)
							marksSelectedPeriodByFilials[dictsOuterKey].MarkBad++;
					} else if (mark.Contains("Средне")) {
						marksByWeeks[dictsOuterKey][dictByWeeksInnerKey].MarkMedium++;

						if (createDate >= dateTimeBegin)
							marksSelectedPeriodByFilials[dictsOuterKey].MarkMedium++;
					} else if (mark.Contains("Хорошо")) {
						marksByWeeks[dictsOuterKey][dictByWeeksInnerKey].MarkGood++;

						if (createDate >= dateTimeBegin)
							marksSelectedPeriodByFilials[dictsOuterKey].MarkGood++;
					} else if (mark.Contains("Дубль")) {
						if (createDate >= dateTimeBegin)
							marksSelectedPeriodByFilials[dictsOuterKey].MarkDuplicate++;
					} else
						Logging.ToLog("Неизвестная оценка - " + mark);
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			}

			uniqueInnerKeys.Sort();

			foreach (string innerKey in uniqueInnerKeys)
				foreach (string outerKey in marksByWeeks.Keys)
					if (!marksByWeeks[outerKey].Keys.Contains(innerKey))
						marksByWeeks[outerKey].Add(innerKey, new ItemRegistryMark("", ""));

			Excel.Worksheet ws = wb.Sheets["Сводная таблица"];
			ws.Activate();
			RegistryMarkDrawPivotTable(ws, marksSelectedPeriodByFilials);

			ws = wb.Sheets["График - кол-во"];
			ws.Activate();
			RegistryMarkDrawMarksByWeek(xlApp, ws, marksByWeeks, uniqueInnerKeys, RegistryMarkChartType.Total);

			//ws = wb.Sheets["График - %"];
			//ws.Activate();
			//RegistryMarkDrawMarksByWeek(xlApp, ws, marksByWeeks, uniqueInnerKeys, 1);

			//ws = wb.Sheets["График - KPI"];
			//ws.Activate();
			//RegistryMarkDrawMarksByWeek(xlApp, ws, marksByWeeks, uniqueInnerKeys, 2);

			Marshal.ReleaseComObject(ws);
			ws = null;
		}

		private enum RegistryMarkChartType {
			Total, Percentage, KPI
		}

		private static void RegistryMarkDrawPivotTable(
			Excel.Worksheet ws, SortedDictionary<string, ItemRegistryMark> marksSelectedPeriodByFilials) {
			int row = 2;
			int markBadTotal = 0;
			int markMediumTotal = 0;
			int markGoogTotal = 0;
			int markDuplicateTotal = 0;

			foreach (ItemRegistryMark item in marksSelectedPeriodByFilials.Values) {
				ws.Range["A" + row].Value = item.FilialName;
				ws.Range["B" + row].Value = item.Department;
				ws.Range["C" + row].Value = item.MarkBad;
				ws.Range["D" + row].Value = item.MarkMedium;
				ws.Range["E" + row].Value = item.MarkGood;
				ws.Range["F" + row].Value = item.MarkTotal;
				ws.Range["G" + row].Value = (item.MarkTotal > 0) ? (double)item.MarkBad / (double)item.MarkTotal : 0;
				ws.Range["H" + row].Value = (item.MarkTotal > 0) ? (double)item.MarkMedium / (double)item.MarkTotal : 0;
				ws.Range["I" + row].Value = (item.MarkTotal > 0) ? (double)item.MarkGood / (double)item.MarkTotal : 0;
				ws.Range["J" + row].Value = item.MarkDuplicate;

				markBadTotal += item.MarkBad;
				markMediumTotal += item.MarkMedium;
				markGoogTotal += item.MarkGood;
				markDuplicateTotal += item.MarkDuplicate;

				row++;
			}

			int totalMarks = markBadTotal + markMediumTotal + markGoogTotal;
			ws.Range["A" + row].Value = "Итого";
			ws.Range["C" + row].Value = markBadTotal;
			ws.Range["D" + row].Value = markMediumTotal;
			ws.Range["E" + row].Value = markGoogTotal;
			ws.Range["F" + row].Value = totalMarks;
			ws.Range["G" + row].Value = (totalMarks > 0) ? (double)markBadTotal / (double)totalMarks : 0;
			ws.Range["H" + row].Value = (totalMarks > 0) ? (double)markMediumTotal / (double)totalMarks : 0;
			ws.Range["I" + row].Value = (totalMarks > 0) ? (double)markGoogTotal / (double)totalMarks : 0;
			ws.Range["J" + row].Value = markDuplicateTotal;

			ws.Columns["G:I"].Style = "Percent";

			AddBoldBorder(ws.Range["A" + row + ":J" + row]);
			AddBoldBorder(ws.Range["A2:B" + row]);
			AddBoldBorder(ws.Range["C2:E" + row]);
			AddBoldBorder(ws.Range["F2:F" + row]);
			AddBoldBorder(ws.Range["G2:I" + row]);
			AddBoldBorder(ws.Range["J2:J" + row]);

			ws.Range["A" + row].Font.Bold = true;

			row += 2;
			ws.Range["A" + row].Value = "* попытки повторного голосования в течении 60 секунд";
			ws.Range["A" + row].Font.Italic = true;
			ws.Range["A1"].Select();
		}

		private static void RegistryMarkDrawMarksByWeek(
			Excel.Application xlApp, Excel.Worksheet ws,
			SortedDictionary<string, SortedDictionary<string, ItemRegistryMark>> marksByWeeks,
			List<string> uniqueInnerKeys,
			RegistryMarkChartType type) {
			string chartTitle;
			string hint;

			switch (type) {
				case RegistryMarkChartType.Total:
					chartTitle = "Всего оценок - хронология";
					hint = "Отображены все оценки 'Плохо' + 'Средне' + 'Хорошо'";
					break;
				case RegistryMarkChartType.Percentage:
					chartTitle = "Соотношение оценок хорошо и плохо";
					hint = "Отображены только оценки 'Хорошо' и 'Плохо'";
					break;
				case RegistryMarkChartType.KPI:
					chartTitle = "KPI - хронология";
					hint = "KPI рассчитывается по формуле: 'Средне' + 'Хорошо' / 'Всего'";
					break;
				default:
					Logging.ToLog("Неизвестный тип оценки - " + type);
					return;
			}

			int row = 1;
			int column = 1;

			if (type != RegistryMarkChartType.Total)
				foreach (string innerKey in uniqueInnerKeys) {
					ws.Cells[1, column].Value2 = innerKey;
					column++;
				}

			row++;

			foreach (KeyValuePair<string, SortedDictionary<string, ItemRegistryMark>> pair in marksByWeeks) {
				string rowTitle = string.Empty;
				ws.Cells[row, 1].Value2 = pair.Key;
				column = 2;

				foreach (KeyValuePair<string, ItemRegistryMark> weekMarks in pair.Value) {
					object value;
					ItemRegistryMark mark = weekMarks.Value;

					switch (type) {
						case RegistryMarkChartType.Total:
							string[] markDate = weekMarks.Key.Split('/');
							string markYear = markDate[0].Trim(' ');
							int weekNumber = Convert.ToInt32(markDate[1].Trim(' '));
							string currentRowTitle = pair.Key + " " + markYear;

							if (string.IsNullOrEmpty(rowTitle)) {
								rowTitle = currentRowTitle;
								ws.Cells[row, 1].Value2 = currentRowTitle;
							} else {
								if (!rowTitle.Equals(currentRowTitle)) {
									row++;
									rowTitle = currentRowTitle;
									ws.Cells[row, 1].Value2 = rowTitle;
								}
							}

							column = weekNumber + 1;
							value = mark.MarkBad + mark.MarkMedium + mark.MarkGood;
							break;
						case RegistryMarkChartType.Percentage:
							value = weekMarks.Value.MarkGood;
							break;
						case RegistryMarkChartType.KPI:
							if (mark.MarkTotal > 0)
								value = ((double)mark.MarkTotal - (double)mark.MarkBad) / (double)mark.MarkTotal;
							else
								value = string.Empty;
							break;
						default:
							continue;
					}

					ws.Cells[row, column].Value2 = value;

					if (type == RegistryMarkChartType.Percentage ||
						type == RegistryMarkChartType.KPI)
						ws.Cells[row, column].NumberFormat = "0%";

					column++;
				}

				row++;
			}

			column = 1;
			ws.Cells[row, column].Value2 = "ИТОГО";
			foreach (string innerKey in uniqueInnerKeys) {
				column++;
				if (type == RegistryMarkChartType.KPI) {
					double marksPositive = 0;
					double marksTotal = 0;
					foreach (KeyValuePair<string, SortedDictionary<String, ItemRegistryMark>> pair in marksByWeeks) {
						ItemRegistryMark mark = pair.Value[innerKey];
						marksPositive += mark.MarkTotal - mark.MarkBad;
						marksTotal += mark.MarkTotal;
					}

					if (marksTotal > 0)
						ws.Cells[row, column].Value2 = marksPositive / marksTotal;
					else
						ws.Cells[row, column].Value2 = string.Empty;

					ws.Cells[row, column].NumberFormat = "0%";
				} else
					ws.Cells[row, column].FormulaR1C1Local = "=СУММ(R[-" + (row - 2) + "]C:R[-1]C)";
			}

			Excel.Shape shape = xlApp.ActiveSheet.Shapes.AddChart2(234, Excel.XlChartType.xlLineMarkers, 10, 200, 1350, 370);
			shape.Select();
			xlApp.ActiveChart.SetSourceData(ws.UsedRange);
			xlApp.ActiveChart.ChartTitle.Text = chartTitle;

			for (int i = 1; i <= marksByWeeks.Keys.Count; i++)
				xlApp.ActiveChart.FullSeriesCollection(i).IsFiltered = true;

			ws.Cells[row + 2, 1].Value2 = hint;
			ws.Cells[row + 2, 1].Font.Italic = true;

			ws.Range["A1"].Select();

			Marshal.ReleaseComObject(shape);
			shape = null;
		}

		private class ItemRegistryMark {
			public string FilialName { get; private set; }
			public string Department { get; private set; }
			public int MarkBad { get; set; }
			public int MarkMedium { get; set; }
			public int MarkGood { get; set; }
			public int MarkDuplicate { get; set; }

			public int MarkTotal { get { return MarkBad + MarkMedium + MarkGood; } }

			public ItemRegistryMark(string filialName, string department) {
				FilialName = filialName;
				Department = department;
			}
		}

	}
}
