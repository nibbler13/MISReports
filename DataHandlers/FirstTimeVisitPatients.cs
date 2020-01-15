using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class FirstTimeVisitPatients : ExcelGeneral {
		public static bool Process(string resultFile, DataTable dataTable) {
			CopyFormatting(resultFile);

			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws))
				return false;

			try {
				ws.UsedRange.Rows.AutoFit();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				int totalPatients = 0;
				SortedDictionary<string, int> results = new SortedDictionary<string, int>();

				foreach (DataRow dataRow in dataTable.Rows) {
					string filial = dataRow[4].ToString();

					if (filial.Equals("Филиал"))
						continue;

					if (string.IsNullOrEmpty(filial))
						filial = "Пусто";

					if (!results.ContainsKey(filial))
						results.Add(filial, 0);

					results[filial]++;
					totalPatients++;
				}

				ws = wb.Sheets["Сводная таблица"];
				ws.Activate();

				ws.Range["A1"].Value2 = "Филиал";
				ws.Range["B1"].Value2 = "Кол-во";

				int row = 2;
				foreach (KeyValuePair<string, int> pair in results) {
					ws.Range["A" + row].Value2 = pair.Key;
					ws.Range["B" + row].Value2 = pair.Value;
					row++;
				}

				ws.Range["A" + row].Value2 = "Общий итог";
				ws.Range["B" + row].Value2 = totalPatients;

				ws.Columns["A:A"].ColumnWidth = 20;
				ws.UsedRange.Select();
				xlApp.Selection.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlDot;
				xlApp.Selection.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlDot;
				xlApp.Selection.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlDot;
				xlApp.Selection.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDot;
				xlApp.Selection.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlDot;
				xlApp.Selection.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlDot;

				ws.Range["A1:B1"].Select();
				AddInteriorColor(ws.Range["A1:B1"], Excel.XlThemeColor.xlThemeColorAccent5);
				AddInteriorColor(ws.Range["A" + row + ":B" + row], Excel.XlThemeColor.xlThemeColorAccent6);

				ws.Range["A1:A" + row].Font.Bold = true;
				ws.Range["A1:B1"].Font.Bold = true;
				ws.Range["A1"].Select();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			wb.Sheets["Сводная таблица"].Activate();
			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}
	}
}
