using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports {
	class NpoiExcelUnclosedProtocols {
		public static string WriteDataTableToExcel(DataTable dataTable, string resultFilePrefix) {
			string templateFile = Program.AssemblyDirectory + "TemplateUnclosedProtocols.xlsx";
			foreach (char item in Path.GetInvalidFileNameChars())
				resultFilePrefix = resultFilePrefix.Replace(item, '-');

			if (!File.Exists(templateFile))
				return "Не удалось найти файл шаблона: " + templateFile;

			string resultPath = Path.Combine(Program.AssemblyDirectory, "Results");
			if (!Directory.Exists(resultPath))
				Directory.CreateDirectory(resultPath);

			string resultFile = Path.Combine(resultPath, resultFilePrefix + ".xlsx");
			
			IWorkbook workbook;
			using (FileStream stream = new FileStream(templateFile, FileMode.Open, FileAccess.Read))
				workbook = new XSSFWorkbook(stream);

			int rowNumber = 1;
			int columnNumber = 0;

			ISheet sheet = workbook.GetSheet("Подробности");

			foreach (DataRow dataRow in dataTable.Rows) {
				IRow row = sheet.CreateRow(rowNumber);

				foreach (DataColumn column in dataTable.Columns) {
					ICell cell = row.CreateCell(columnNumber);
					string value = dataRow[column].ToString().Replace(" 0:00:00", "");

					if (double.TryParse(value, out double result))
						cell.SetCellValue(result);
					else
						cell.SetCellValue(value);

					columnNumber++;
				}

				columnNumber = 0;
				rowNumber++;
			}

			using (FileStream stream = new FileStream(resultFile, FileMode.Create, FileAccess.Write))
				workbook.Write(stream);

			workbook.Close();

			Excel.Application xlApp = new Excel.Application();

			if (xlApp == null)
				return "Не удалось открыть приложение Excel";

			xlApp.Visible = false;

			Excel.Workbook wb = xlApp.Workbooks.Open(resultFile);

			if (wb == null)
				return "Не удалось открыть книгу " + resultFile;

			Excel.Worksheet ws = wb.Sheets["Подробности"];

			if (ws == null)
				return "Не удалось открыть лист Подробности";

			try {
				PerformSheet(wb, ws, xlApp);
			} catch (Exception e) {
				SystemLogging.LogMessageToFile(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				AddPivotTableDepartments(wb, ws, xlApp);
			} catch (Exception e) {
				SystemLogging.LogMessageToFile(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				AddPivotTableDoctors(wb, ws, xlApp);
			} catch (Exception e) {
				SystemLogging.LogMessageToFile(e.Message + Environment.NewLine + e.StackTrace);
			}

			wb.Sheets["Сводная по врачам"].Activate();

			wb.Save();
			wb.Close();

			xlApp.Quit();

			return resultFile;
		}

		private static void PerformSheet(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			int usedRows = ws.UsedRange.Rows.Count;
			ws.Range["A1"].Select();
			xlApp.Selection.AutoFilter();
			ws.Columns["B:B"].Select();
			xlApp.Selection.Insert(Excel.XlDirection.xlToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
			ws.Range["B1"].Value = "Уникальное лечение";
			ws.Range["B2"].Select();
			xlApp.ActiveCell.FormulaR1C1 = "=IF(RC[-1]=R[1]C[-1],0,1)";
			xlApp.Selection.AutoFill(ws.Range["B2:B" + usedRows]);

		}

			private static void AddPivotTableDoctors(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			ws.Cells[1, 1].Select();

			string pivotTableName = @"WorkTimePivotTable";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная по врачам"];

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			pivotTable.PivotFields("ФИО врача").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ФИО врача").Position = 1;

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 2;

			pivotTable.PivotFields("Подразделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Подразделение").Position = 3;

			pivotTable.PivotFields("ФИО врача").Subtotals = 
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("ФИО врача").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Филиал").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Филиал").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.PivotFields("Подразделение").Subtotals =
				new bool[] { false, false, false, false, false, false, false, false, false, false, false, false };
			pivotTable.PivotFields("Подразделение").LayoutForm = Excel.XlLayoutFormType.xlTabular;

			pivotTable.AddDataField(pivotTable.PivotFields("Уникальное лечение"), "Кол-во лечений", Excel.XlConsolidationFunction.xlSum);

			pivotTable.CalculatedFields().Add("Кол-во протоколов", "='Протокол подписан' +'Протокол не подписан'", true);
			pivotTable.PivotFields("Кол-во протоколов").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю Кол-во протоколов").Caption = "Общее кол-во протоколов";

			pivotTable.AddDataField(pivotTable.PivotFields("Протокол не подписан"), "Кол-во неподписанных", Excel.XlConsolidationFunction.xlSum);

			pivotTable.CalculatedFields().Add("Процент неподписанных", "='Протокол не подписан' /'Кол-во протоколов'", Excel.XlConsolidationFunction.xlSum);
			pivotTable.PivotFields("Процент неподписанных").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю Процент неподписанных").Caption = "Доля неподписанных протоколов";
			pivotTable.PivotFields("Доля неподписанных протоколов").NumberFormat = "0,00%";
			
			pivotTable.PivotFields("ФИО врача").AutoSort(Excel.XlSortOrder.xlDescending,
				"Доля неподписанных протоколов");

			pivotTable.HasAutoFormat = false;
			wb.ShowPivotTableFieldList = false;
		}

		private static void AddPivotTableDepartments(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			ws.Cells[1, 1].Select();

			string pivotTableName = @"WorkTimePivotTable";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная по отделениям"];

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);

			pivotTable.PivotFields("Филиал").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Филиал").Position = 1;

			pivotTable.PivotFields("Подразделение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Подразделение").Position = 2;

			pivotTable.PivotFields("ФИО врача").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ФИО врача").Position = 3;

			pivotTable.AddDataField(pivotTable.PivotFields("Уникальное лечение"), "Кол-во лечений", Excel.XlConsolidationFunction.xlSum);

			pivotTable.CalculatedFields().Add("Всего протоколов", "='Протокол подписан' +'Протокол не подписан'", true);
			pivotTable.PivotFields("Всего протоколов").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю Всего протоколов").Caption = "Общее кол-во протоколов";

			pivotTable.AddDataField(pivotTable.PivotFields("Протокол не подписан"), "Кол-во неподписанных протоколов", Excel.XlConsolidationFunction.xlSum);

			pivotTable.CalculatedFields().Add("Процент неподписанных", "='Протокол не подписан' /'Всего протоколов'", true);
			pivotTable.PivotFields("Процент неподписанных").Orientation = Excel.XlPivotFieldOrientation.xlDataField;
			pivotTable.PivotFields("Сумма по полю Процент неподписанных").Caption = "Доля неподписанных протоколов";
			pivotTable.PivotFields("Доля неподписанных протоколов").NumberFormat = "0,00%";

			pivotTable.PivotFields("Филиал").AutoSort(Excel.XlSortOrder.xlDescending, "Доля неподписанных протоколов");
			pivotTable.PivotFields("Подразделение").AutoSort(Excel.XlSortOrder.xlDescending, "Доля неподписанных протоколов");
			pivotTable.PivotFields("ФИО врача").AutoSort(Excel.XlSortOrder.xlDescending, "Доля неподписанных протоколов");

			pivotTable.PivotFields("Подразделение").ShowDetail = false;
			pivotTable.PivotFields("Филиал").ShowDetail = false;

			pivotTable.HasAutoFormat = false;

			wsPivote.Columns[1].ColumnWidth = 60;
			wb.ShowPivotTableFieldList = false;
		}
	}
}
