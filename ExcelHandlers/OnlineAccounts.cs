using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class OnlineAccounts : ExcelGeneral {

		//============================ OnlineAccountsUsage ============================
		public static bool Process(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws))
				return false;

			try {
				int rowsUsed = ws.UsedRange.Rows.Count;

				for (int i = 2; i <= rowsUsed; i++)
					ws.Range["F" + i].FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)";

				ws.Columns["F:F"].Select();
				xlApp.Selection.NumberFormat = "0,0%";

				ws.Range["A" + (rowsUsed + 2)].Value = "Итого:";

				foreach (string item in new string[] { "B", "C", "D", "E" })
					ws.Range[item + (rowsUsed + 2)].Formula = "=SUM(" + item + "2:" + item + rowsUsed + ")";

				ws.Range["F" + (rowsUsed + 2)].FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)";


				//string rangeData = "A1:A" + rowsUsed + ",F1:F" + rowsUsed;
				//Console.WriteLine("rangeData: " + rangeData);
				xlApp.ActiveSheet.Shapes.AddChart2(201, Excel.XlChartType.xlColumnClustered).Select();
				xlApp.ActiveChart.SetSourceData(ws.get_Range("A1:A2;F1:F2"));
				xlApp.ActiveSheet.Shapes["Диаграмма 1"].Top = 0;
				xlApp.ActiveSheet.Shapes["Диаграмма 1"].Left = 480;

				//rowsUsed += 2;
				//ws.Range["A" + rowsUsed].Value = "Итого:";
				//ws.Range["B" + rowsUsed].Formula = "=AVERAGE(B2:B" + (rowsUsed - 2) + ")";
				//ws.Range["A" + rowsUsed + ":B" + rowsUsed].Select();
				//xlApp.Selection.Interior.Pattern = Excel.Constants.xlSolid;
				//xlApp.Selection.Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
				//xlApp.Selection.Interior.Color = 65535;
				//xlApp.Selection.Interior.TintAndShade = 0;
				//xlApp.Selection.Font.Bold = Excel.Constants.xlSolid;
				//rowsUsed++;
				ws.Range["A" + rowsUsed].Select();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}


	}
}
