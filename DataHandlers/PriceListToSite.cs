using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class PriceListToSite : ExcelGeneral {
		public static DataTable PerformData(DataTable main, DataTable exclusions, DataTable grouping, DataTable priorities, out string emptyFields) {
			DataTable dataTableResult = main.Clone();

			emptyFields = string.Empty;

			foreach (DataRow dataRow in main.Rows) {
				if (IsPriceTooLow(dataRow))
					continue;

				dataRow["PRIORITY"] = GetPriority(dataRow, priorities);
				dataRow["TYPE_NAME"] = GetType(dataRow);
				GetTopLevelAndSiteService(dataRow, grouping, out string topLevel, out string siteService);
				dataRow["TOP_LEVEL"] = topLevel;
				dataRow["SITE_SERVICE"] = siteService;

				if (Debugger.IsAttached)
					if (dataRow["SUBGROUP"].ToString().Equals("ТЕСТЫ ДЛЯ ПРОФОСМОТРОВ, ВЫПОЛНЯЕМЫЕ В МОСКВЕ И РЕГИОНАХ ПО ДОГОВОРАМ"))
						Console.WriteLine("");

				if (IsNeedToExclude(dataRow, exclusions)) {
					Console.WriteLine("Пропуск строки: " + string.Join("; ", dataRow.ItemArray));
					continue;
				}

				string serviceHeaders = dataRow["TOP_LEVEL"].ToString() + " " + dataRow["GROUP_NAME"].ToString() + " " + dataRow["SUBGROUP"].ToString();
				serviceHeaders = serviceHeaders.ToLower();

				if (serviceHeaders.Contains("детск") ||
					serviceHeaders.Contains("дети") ||
					serviceHeaders.Contains("детей") ||
					serviceHeaders.Contains("детям")) {
					try {
						dataRow["KID_MSSU_NAL"] = dataRow["ADULT_MSSU_NAL"];
						dataRow["ADULT_MSSU_NAL"] = DBNull.Value;
					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					}
				} else if (topLevel.Equals("Лабораторные исследования")) {
					try {
						dataRow["Zabornik1_NAL"] = dataRow["MSPO_NAL"];
						dataRow["Zabornik2_NAL"] = dataRow["MSPO_NAL"];
					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					}
				}

                if (string.IsNullOrEmpty(topLevel) ||
                    string.IsNullOrEmpty(siteService))
                    emptyFields += "<tr><td>" + string.Join("</td><td>", dataRow.ItemArray) + "</td></tr>";

				dataTableResult.Rows.Add(dataRow.ItemArray);
			}

            if (!string.IsNullOrEmpty(emptyFields))
                emptyFields =
                    "<table border='1'><tr><th>ВЕРХНИЙ УРОВЕНЬ</th><th>ГРУППА</th><th>ПОДГРУППА</th><th>УСЛУГА САЙТА</th><th>ID УСЛУГИ</th><th>ВИД</th><th>ПРИОРИТЕТ</th><th>" +
					"КОД УСЛУГИ</th><th>ИМЯ УСЛУГИ</th><th>МДМ_Наличный расчет</th><th>Взросл_СУЩ_Наличный расчет</th><th>Детск_СУЩ_Наличный расчет</th><th>М-СРЕТ_Наличный расчет</th><th>" +
                    "С-Пб._Наличный расчет</th><th>Красн_Наличный расчет</th><th>Уфа_Наличный расчет</th><th>Казань_Наличный расчет</th><th>" +
					"К-УРАЛ_Наличный расчет</th><th>Сочи_Наличный расчет</th><th>Заборник_1_Наличный расчет</th><th>Заборник_2_Наличный расчет</th></tr>" + emptyFields + "</table>";

			return dataTableResult;
		}

		private static bool IsPriceTooLow(DataRow dataRow) {
			bool hasPriceMoreThan2 = false;

			try {
				for (int i = 9; i <= 17; i++)
					if (double.TryParse(dataRow[i].ToString(), out double price))
						if (price > 2) {
							hasPriceMoreThan2 = true;
							break;
						}
			} catch (Exception e) {
				Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
			}

			return !hasPriceMoreThan2;
		}

		private static bool IsNeedToExclude(DataRow dataRow, DataTable exclusions) {
			bool isNeedToExclude = false;

			if (exclusions == null)
				return isNeedToExclude;

			try {
				string[] typesToExclude = new string[] { "GROUP_NAME", "SUBGROUP", "KODOPER" };
				foreach (string excludeType in typesToExclude) {
					string valueToCheck = dataRow[excludeType].ToString();

					if (string.IsNullOrEmpty(valueToCheck) ||
						string.IsNullOrWhiteSpace(valueToCheck))
						continue;

					EnumerableRowCollection<DataRow> exclusionsSearchGroup =
						from row in exclusions.AsEnumerable()
						where row.Field<string>(1) != null && row.Field<string>(1).ToLower().Equals(valueToCheck.ToLower())
						select row;

					if (exclusionsSearchGroup != null && exclusionsSearchGroup.Count() > 0)
						foreach (DataRow excludeRow in exclusionsSearchGroup)
							if (excludeRow[0].ToString().ToLower().Equals(excludeType.ToLower())) {
								Console.WriteLine("Excluded: " + valueToCheck);
								isNeedToExclude = true;
								break;
							}
				}
			} catch (Exception e) {
				Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
			}

			return isNeedToExclude;
		}

		private static int GetPriority(DataRow dataRow, DataTable priorities) {
			int priority = 99;
			string kodoper = dataRow["KODOPER"].ToString();

			try {
				if (string.IsNullOrEmpty(kodoper))
					return priority;

				EnumerableRowCollection<DataRow> rows =
					from row in priorities.AsEnumerable()
					where row.Field<string>(7) != null && row.Field<string>(7).Equals(kodoper)
					select row;

				if (rows.Count() == 1) {
					DataRow row = rows.First();
					if (int.TryParse(row[6].ToString(), out int parsedPriority))
						priority = parsedPriority;
				}
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			return priority;
		}

		private static string GetType(DataRow dataRow) {
			string type = "Услуги";

			try {
				string serviceName = dataRow["SERVICE_NAME"].ToString();
				serviceName = serviceName.ToLower();
				if (serviceName.Contains("прием") ||
					serviceName.Contains("приём") ||
					serviceName.Contains("консультация"))
					type = "Прием врача";

			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			return type;
		}

		private static void GetTopLevelAndSiteService(DataRow dataRow, DataTable grouping, out string topLevel, out string siteService) {
			topLevel = string.Empty;
			siteService = string.Empty;

			try {
				string kodoper = dataRow["KODOPER"].ToString();
				string group = dataRow["GROUP_NAME"].ToString().ToLower();
				string subgroup = dataRow["SUBGROUP"].ToString().ToLower();

				EnumerableRowCollection<DataRow> rowsByKodoper =
					from row in grouping.AsEnumerable()
					where (row.Field<string>(0) != null && row.Field<string>(0).Equals(kodoper))
					select row;

				if (rowsByKodoper.Count() >= 1) {
					DataRow dataRowGrouping = rowsByKodoper.First();
					SetupTopLevelAndSiteServiceValue(dataRowGrouping, ref topLevel, ref siteService);
					return;
				}

				string type = dataRow["TYPE_NAME"].ToString().ToLower();

				EnumerableRowCollection<DataRow> rowsByGroup =
					from row in grouping.AsEnumerable()
					where (row.Field<string>(1) != null && row.Field<string>(1).ToLower().Equals(@group))
					select row;

				if (IsEnumerableCollectionGroupFounded(rowsByGroup, type, ref topLevel, ref siteService))
					return;

				EnumerableRowCollection<DataRow> rowsBySubgroup =
					from row in grouping.AsEnumerable()
					where (row.Field<string>(2) != null && row.Field<string>(2).ToLower().Equals(subgroup))
					select row;

				IsEnumerableCollectionGroupFounded(rowsBySubgroup, type, ref topLevel, ref siteService);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}
		}

		private static void SetupTopLevelAndSiteServiceValue(DataRow dataRow, ref string topLevel, ref string siteService) {
			topLevel = dataRow[5].ToString();
			siteService = dataRow[4].ToString();
		}

		private static bool IsEnumerableCollectionGroupFounded(EnumerableRowCollection<DataRow> collection,
														 string type,
														 ref string topLevel,
														 ref string siteService) {
			int rowsBySubgroupCount = collection.Count();

			if (rowsBySubgroupCount == 1) {
				DataRow dataRowBySubgroup = collection.First();
				string typeByGroup = dataRowBySubgroup[3].ToString();
				if (string.IsNullOrEmpty(typeByGroup) ||
					typeByGroup.ToLower().Equals(type)) {
					SetupTopLevelAndSiteServiceValue(dataRowBySubgroup, ref topLevel, ref siteService);
					return true;
				}
			} else if (rowsBySubgroupCount == 2) {
				DataRow dataRowBySubgroupFirst = collection.First();
				DataRow dataRowBySubgroupLast = collection.Last();
				string typeBySubgroupFirst = dataRowBySubgroupFirst[3].ToString();
				string typeBySubgroupLast = dataRowBySubgroupLast[3].ToString();

				if (typeBySubgroupFirst.ToLower().Equals(type)) {
					SetupTopLevelAndSiteServiceValue(dataRowBySubgroupFirst, ref topLevel, ref siteService);
					return true;
				} else if (typeBySubgroupLast.ToLower().Equals(type)) {
					SetupTopLevelAndSiteServiceValue(dataRowBySubgroupLast, ref topLevel, ref siteService);
					return true;
				} else if (string.IsNullOrEmpty(typeBySubgroupFirst)) {
					SetupTopLevelAndSiteServiceValue(dataRowBySubgroupFirst, ref topLevel, ref siteService);
					return true;
				} else if (string.IsNullOrEmpty(typeBySubgroupLast)) {
					SetupTopLevelAndSiteServiceValue(dataRowBySubgroupLast, ref topLevel, ref siteService);
					return true;
				}
			} else if (rowsBySubgroupCount > 2) {
				Logging.ToLog("Результат поиска группировки по подгруппе более 3, пропуск");
			}

			return false;
		}


		public static bool Process(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws))
				return false;

			int usedRows = ws.UsedRange.Rows.Count;

			ws.Range["A3:U3"].Select();
			xlApp.Selection.Copy();
			ws.Range["A4:U" + usedRows].Select();
			xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
			ws.Range["A1"].Select();

			SaveAndCloseWorkbook(xlApp, wb, ws);

            return true;
		}
    }
}
