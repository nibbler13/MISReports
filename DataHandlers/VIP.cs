using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class VIP : ExcelGeneral {

		//============================ VIP ============================
		public static bool Process(string resultFile, string previousFile) {
			Logging.ToLog("Подготовка файла с отчетом по VIP-пациентам: " + resultFile);
			Logging.ToLog("Предыдущий файл: " + previousFile);
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws)) {
				Logging.ToLog("Не удалось открыть книгу: " + resultFile);
				return false;
			}

			try {
				ws.Columns["B:B"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ";
				ws.Columns["C:D"].Select();
				xlApp.Selection.NumberFormat = "ч:мм;@";
				ws.Columns["L:L"].Select();
				xlApp.Selection.NumberFormat = "ДД.ММ.ГГГГ";
				ws.Cells[1, 1].Select();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			if (string.IsNullOrEmpty(previousFile) || !File.Exists(previousFile)) {
				Logging.ToLog("Пропуск сравнения с предыдущей версией, файл не существует");
				return true;
			}

			Logging.ToLog("Считывание содержимого файлов");

			DataTable dataTableCurrent = ReadExcelFile(resultFile, "Данные");
			Logging.ToLog("Текущий файл, строк: " + dataTableCurrent.Rows.Count);

			DataTable dataTablePrevious = ReadExcelFile(previousFile, "Данные");
			Logging.ToLog("Предыдущий файл, строк: " + dataTablePrevious.Rows.Count);

			if (dataTablePrevious.Columns.Count == 15)
				dataTablePrevious.Columns.RemoveAt(14);

			if (!OpenWorkbook(resultFile, out xlApp, out wb, out ws)) {
				Logging.ToLog("Не удалось открыть книгу: " + resultFile);
				return false;
			}

			for (int i = 1; i < dataTableCurrent.Rows.Count; i++) {
				DataRow dataRowLeft = dataTableCurrent.Rows[i];
				bool existedBefore = false;

				for (int k = 1; k < dataTablePrevious.Rows.Count; k++) {
					DataRow dataRowRight = dataTablePrevious.Rows[k];
					if (DataRowComparer.Default.Equals(dataRowLeft, dataRowRight)) {
						existedBefore = true;
						break;
					}
				}

				if (!existedBefore) {
					int rowNumber = i + 1;
					ws.Range["A" + rowNumber + ":O" + rowNumber].Interior.ColorIndex = 35;
					ws.Range["O" + rowNumber].Value2 = "Новая запись";
				}
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

	}
}
