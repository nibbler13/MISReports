using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class PriceListToSite : ExcelGeneral {
		public static DataTable PerformData(DataTable main, DataTable exclusions, DataTable grouping, DataTable priorities) {
			DataTable dataTableResult = main.Clone();

			foreach (DataRow dataRow in main.Rows) {
				if (IsPriceTooLow(dataRow))
					continue;

				if (IsNeedToExclude(dataRow, exclusions))
					continue;

				dataRow["ПРИОРИТЕТ"] = GetPriority(dataRow, priorities);
				dataRow["ВИД"] = GetType(dataRow);
				GetTopLevelAndSiteService(dataRow, grouping, out string topLevel, out string siteService);
				dataRow["ВЕРХНИЙ УРОВЕНЬ"] = topLevel;
				dataRow["УСЛУГА САЙТА"] = siteService;

				dataTableResult.Rows.Add(dataRow.ItemArray);
			}

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

			if (exclusions != null) {
				try {
					string[] typesToExclude = new string[] { "ГРУППА", "ПОДГРУППА" };
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
									isNeedToExclude = true;
									break;
								}
					}
				} catch (Exception e) {
					Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
				}
			}

			return isNeedToExclude;
		}

		private static int GetPriority(DataRow dataRow, DataTable priorities) {
			int priority = 99;
			string kodoper = dataRow["КОД УСЛУГИ"].ToString();

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
				string serviceName = dataRow["ИМЯ УСЛУГИ"].ToString();
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
				string kodoper = dataRow["КОД УСЛУГИ"].ToString();
				string group = dataRow["ГРУППА"].ToString().ToLower();
				string subgroup = dataRow["ПОДГРУППА"].ToString().ToLower();

				EnumerableRowCollection<DataRow> rowsByKodoper =
					from row in grouping.AsEnumerable()
					where (row.Field<string>(0) != null && row.Field<string>(0).Equals(kodoper))
					select row;

				if (rowsByKodoper.Count() >= 1) {
					DataRow dataRowGrouping = rowsByKodoper.First();
					SetupTopLevelAndSiteServiceValue(dataRowGrouping, ref topLevel, ref siteService);
					return;
				}

				string type = dataRow["ВИД"].ToString().ToLower();

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

			ws.Range["A3:R3"].Select();
			xlApp.Selection.Copy();
			ws.Range["A4:R" + usedRows].Select();
			xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
			ws.Range["A1"].Select();

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}
	}
}
