using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports {
	class NpoiExcelMESUsage {
		public static string WriteTreatmentsToExcel(Dictionary<string, MESUsage.ItemTreatment> treatments, string resultFilePrefix) {
			string templateFile = Program.AssemblyDirectory + "TemplateMESUsage.xlsx";
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

			foreach (KeyValuePair<string, MESUsage.ItemTreatment> treatment in treatments) {
				IRow row = sheet.CreateRow(rowNumber);

				MESUsage.ItemTreatment treat = treatment.Value;
				double percentCompleted = 
					((double)treat.ListReferralsFromMes.Count + 
					(double)treat.ListReferralsFromDoc.Count) / 
					(double)treat.ListMES.Count;

				List<object> values = new List<object>() {
					treatment.Key,
					treat.TREATDATE,
					treat.FILIAL,
					treat.DEPNAME,
					treat.DOCNAME,
					treat.DOCPOSITION,
					treat.HISTNUM,
					treat.CLIENTNAME,
					treat.AGE,
					treat.MKBCODE,
					treat.ListMES.Count,
					treat.ListReferralsFromMes.Count > 0 ? 1 : 0,
					treat.ListReferralsFromMes.Count,
					treat.ListReferralsFromDoc.Count,
					percentCompleted,
					percentCompleted == 1 ? 1 : 0
				};

				foreach (object value in values) {
					ICell cell = row.CreateCell(columnNumber);

					if (double.TryParse(value.ToString(), out double result))
						cell.SetCellValue(result);
					else
						cell.SetCellValue(value.ToString());

					columnNumber++;
				}

				columnNumber = 0;
				rowNumber++;
			}

			using (FileStream stream = new FileStream(resultFile, FileMode.Create, FileAccess.Write))
				workbook.Write(stream);

			workbook.Close();

			//Excel.Application xlApp = new Excel.Application();

			//if (xlApp == null)
			//	return "Не удалось открыть приложение Excel";

			//xlApp.Visible = false;

			//Excel.Workbook wb = xlApp.Workbooks.Open(resultFile);

			//if (wb == null)
			//	return "Не удалось открыть книгу " + resultFile;

			//Excel.Worksheet ws = wb.Sheets["Подробности"];

			//if (ws == null)
			//	return "Не удалось открыть лист Подробности";

			//try {
			//	PerformSheet(wb, ws, xlApp);
			//} catch (Exception e) {
			//	SystemLogging.LogMessageToFile(e.Message + Environment.NewLine + e.StackTrace);
			//}

			//try {
			//	AddPivotTableDepartments(wb, ws, xlApp);
			//} catch (Exception e) {
			//	SystemLogging.LogMessageToFile(e.Message + Environment.NewLine + e.StackTrace);
			//}

			//try {
			//	AddPivotTableDoctors(wb, ws, xlApp);
			//} catch (Exception e) {
			//	SystemLogging.LogMessageToFile(e.Message + Environment.NewLine + e.StackTrace);
			//}

			//wb.Sheets["Сводная по врачам"].Activate();

			//wb.Save();
			//wb.Close();

			//xlApp.Quit();

			return resultFile;
		}
	}
}
