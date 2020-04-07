using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class RegistryMotivation : ExcelGeneral {
		public static bool Process(string resultFile) {
			Logging.ToLog("Выполнение пост-обработки");
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws))
				return false;

			try {
				ws.UsedRange.Select();
				xlApp.Selection.AutoFilter();

				ws.Range["G2:Q2"].Select();
				xlApp.Selection.AutoFill(ws.Range["G2:Q" + ws.UsedRange.Rows.Count]);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			ws.Range["A1"].Select();
			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}
	}
}
