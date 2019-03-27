using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class GBooking : ExcelGeneral {

		//============================ GBooking ============================
		public static bool Process(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws))
				return false;

			try {
				int rowsUsed = ws.UsedRange.Rows.Count;

				ws.Range["A2:I2"].Select();
				xlApp.Selection.Copy();
				ws.Range["A3:I" + rowsUsed].Select();
				xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
				ws.Rows["2:" + rowsUsed].Select();
				xlApp.Selection.RowHeight = 15;

				ws.Range["A1"].Select();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}
	}
}
