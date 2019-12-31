using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class UniqueServices : ExcelGeneral {

		//============================ UniqueServices ============================
		public static string Process(DataTable dataTableCurrent,
											 DataTable dataTableTotal,
											 DataTable dataTableLab,
											 DataTable dataTableLabTotal,
											 string resultFilePrefix,
											 string templateName,
											 string period,
											 ReportsInfo.Type reportType) {
			if (!GetTemplateFilePath(ref templateName))
				return string.Empty;

			string resultPath = GetResultFilePath(resultFilePrefix, templateName);

			try {
				File.Copy(templateName, resultPath);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				return string.Empty;
			}

			if (!OpenWorkbook(resultPath, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws))
				return string.Empty;

			Dictionary<string, int> serviceMap = new Dictionary<string, int> {
				{ "Имплантация (кол-во имплантов)", 7},
				{ "Протезирование: МК, включая коронки на имплантатах", 8},
				{ "Отбеливание ZOOM", 9},
				{ "Ортодонтия: кол-во начатых лечений (уник. обратившихся) в отчетный период", 10},
				{ "Направление на КЛКТ", 11},
				{ "ЭГДС под внутренней седацией", 12},
				{ "ФКС под внутренней седацией", 13},
				{ "Консультация диетолога", 14},
				{ "Закрытые направления КДЛ за наличный расчет (кол-во шт.)", 15},
				{ "Ударно-волновая терапия", 16 }
			};

			Dictionary<string, string> filialMapCurrent = new Dictionary<string, string> {
				{ "МДМ", "G" },
				{ "СУЩ", "H" },
				{ "М-СРЕТ", "I" }
			};

			Dictionary<string, string> filialMapTotal = new Dictionary<string, string> {
				{ "МДМ", "K" },
				{ "СУЩ", "L" },
				{ "М-СРЕТ", "M" }
			};

			Dictionary<string, string> filialMapPlan = new Dictionary<string, string> {
				{ "МДМ", "C" },
				{ "СУЩ", "D" },
				{ "М-СРЕТ", "E" }
			};

			if (reportType == ReportsInfo.Type.UniqueServicesRegions) {
				serviceMap = new Dictionary<string, int> {
					{ "Имплантация (кол-во имплантов)", 7 },
					{ "Протезирование: МК, включая коронки на имплантатах", 8 },
					{ "Ударно-волновая терапия (травматология-ортопедия)", 9 },
					{ "Консультация диетолога (первичная)", 10 },
					{ "Количество уникально обратившихся за наличный расчет гинекология", 11 },
					{ "Количество уникально обратившихся за наличный расчет стоматология", 12 },
					{ "Количество уникально обратившихся за наличный расчет урология", 13 },
					{ "Закрытые направления КДЛ за наличный расчет (кол-во шт.)", 14 }
				};

				filialMapCurrent = new Dictionary<string, string> {
					{ "С-Пб.", "J" },
					{ "Уфа", "K" },
					{ "К-УРАЛ", "L" },
					{ "Казань", "M" },
					{ "Красн", "N" },
					{ "Сочи", "O" }
				};

				filialMapTotal = new Dictionary<string, string> {
					{ "С-Пб.", "Q" },
					{ "Уфа", "R" },
					{ "К-УРАЛ", "S" },
					{ "Казань", "T" },
					{ "Красн", "U" },
					{ "Сочи", "V" }
				};

				filialMapPlan = new Dictionary<string, string> {
					{ "С-Пб.", "C" },
					{ "Уфа", "D" },
					{ "К-УРАЛ", "E" },
					{ "Казань", "F" },
					{ "Красн", "G" },
					{ "Сочи", "H" }
				};
			}

			ParseAndWriteUniqueService(ws, dataTableCurrent, serviceMap, filialMapCurrent, filialMapPlan);
			ParseAndWriteUniqueService(ws, dataTableLab, serviceMap, filialMapCurrent, filialMapPlan);
			ParseAndWriteUniqueService(ws, dataTableTotal, serviceMap, filialMapTotal, filialMapPlan);
			ParseAndWriteUniqueService(ws, dataTableLabTotal, serviceMap, filialMapTotal, filialMapPlan);

			ws.Range["A1"].Value2 = ((string)ws.Range["A1"].Value2).Replace("@period", period);

			string secongDateRange = "G";
			if (reportType == ReportsInfo.Type.UniqueServicesRegions)
				secongDateRange = "J";

			ws.Range[secongDateRange + "5"].Value2 = ((string)ws.Range[secongDateRange + "5"].Value2).Replace("@period", period);

			SaveAndCloseWorkbook(xlApp, wb, ws);
			return resultPath;
		}

		private static void ParseAndWriteUniqueService(Excel.Worksheet ws,
										  DataTable services,
										  Dictionary<string, int> serviceMap,
										  Dictionary<string, string> filialMap,
										  Dictionary<string, string> filialMapPlan) {
			foreach (DataRow dataRow in services.Rows) {
				try {
					string filial = dataRow["SHORTNAME"].ToString().TrimStart(' ').TrimEnd(' ');
					string service = dataRow["SERVICE"].ToString().TrimStart(' ').TrimEnd(' ');
					int scount = Convert.ToInt32(dataRow["SCOUNT"].ToString().TrimStart(' ').TrimEnd(' '));

					if (!serviceMap.Keys.Contains(service) ||
						!filialMap.Keys.Contains(filial) ||
						!filialMapPlan.Keys.Contains(filial)) {
						Logging.ToLog("Не удалось найти ключи для пары: " + filial + "|" + service);
						continue;
					}

					var planValue = ws.Range[filialMapPlan[filial] + serviceMap[service]].Value2;
					if (planValue == null || string.IsNullOrEmpty(planValue.ToString()))
						continue;

					ws.Range[filialMap[filial] + serviceMap[service]].Value2 = scount;
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			}
		}

	}
}
