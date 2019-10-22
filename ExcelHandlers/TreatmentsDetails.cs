using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports.ExcelHandlers {
	class TreatmentsDetails : ExcelGeneral {
		public static void PerformDataTable(ref DataTable dataTable) {
			if (dataTable == null)
				return;

			List<string> doubles = new List<string>();
			string doublesPath = @"\\mskv-fs-02\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\СК РЕСО\Дубли.xlsx";
			Logging.ToLog("Считывание файла с дублями: " + doublesPath);
			if (File.Exists(doublesPath)) {
				try {
					DataTable dataTableDoubles = ReadExcelFile(doublesPath, "Лист1");
					foreach (DataRow row in dataTableDoubles.Rows) {
						string kodoper = row[0].ToString();
						if (string.IsNullOrEmpty(kodoper))
							continue;

						doubles.Add(kodoper);
					}

					Logging.ToLog("Считано кодов услуг: " + doubles.Count);
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			} else {
				Logging.ToLog("Не удается найти (получить доступ) файл с дублями: " + doublesPath);
			}

			List<string> mkbUninsured = new List<string>();
			string mkbUninsuredPath = @"\\mskv-fs-02\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\СК РЕСО\Нестраховые.xlsx";
			Logging.ToLog("Считывание файла с нестраховыми диагнозами: " + mkbUninsuredPath);
			if (File.Exists(mkbUninsuredPath)) {
				try {
					DataTable dataTableUninsured = ReadExcelFile(mkbUninsuredPath, "Лист1");
					foreach (DataRow row in dataTableUninsured.Rows) {
						string mkb = row[0].ToString();
						if (string.IsNullOrEmpty(mkb))
							continue;

						mkbUninsured.Add(mkb);
					}

					Logging.ToLog("Считано кодов МКБ: " + mkbUninsured.Count);
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			} else {
				Logging.ToLog("Не удается найти (получить доступ) файл с нестраховыми услугами: " + mkbUninsuredPath);
			}

			for (int i = 0; i < dataTable.Rows.Count; i++) {
				try {
					DataRow row = dataTable.Rows[i];
					string comment_3 = row["COMMENT_3"].ToString();
					if (!string.IsNullOrEmpty(comment_3))
						continue;

					string amountrub_a = row["AMOUNTRUB_A"].ToString();
					if (string.IsNullOrEmpty(amountrub_a) || (double.TryParse(amountrub_a, out double serviceCost) && serviceCost == 0)) {
						row["COMMENT_3"] = "Нулевые";
						continue;
					}

					string programType = row["PRG"].ToString().ToLower();
					if (programType.Contains("гарантийное письмо")) {
						row["COMMENT_3"] = "ГП";
						continue;
					}

					if (programType.StartsWith("а") || programType.Contains("аванс")) {
						row["COMMENT_3"] = "Аванс";
						continue;
					}

					string age = row["AGE"].ToString();
					if (double.TryParse(age, out double ageParsed) && ageParsed < 1) {
						row["COMMENT_3"] = "Дети";
						continue;
					}

					if(programType.Contains("франшиза")) {
						row["COMMENT_3"] = "Франшиза";
						continue;
					}

					if (programType.Contains("вип") || programType.Contains("vip")) {
						row["COMMENT_3"] = "ВИП";
						continue;
					}

					if (programType.Contains("берем")) {
						row["COMMENT_3"] = "Беременность_программы";
						continue;
					}

					string serviceKodoper = row["KODOPER"].ToString();
					List<string> servicesCodesPregnant = new List<string> {
						"134006", "134008", "134009", "134056", "101205", "101206", "101207", "101908", "134120", "322109", "322114", "329009"
					};

					if (servicesCodesPregnant.Contains(serviceKodoper)) {
						row["COMMENT_3"] = "Беременность_услуги";
						continue;
					}

					List<string> servicesCodeVaccine = new List<string>() {
						"140011", "225004"
					};

					if (servicesCodeVaccine.Contains(serviceKodoper)) {
						row["COMMENT_3"] = "Вакцинация";
						continue;
					}

					List<string> servicesCodesDroppers = new List<string>() {
						"101821", "101884", "155001", "212022"
					};

					if (servicesCodesDroppers.Contains(serviceKodoper)) {
						row["COMMENT_3"] = "Капельницы";
						continue;
					}

					if (doubles.Contains(serviceKodoper)) {
						string treatcode = row["TREATCODE"].ToString();
						bool isDoubled = false;
						for (int x = i + 1; x < dataTable.Rows.Count; x++) {
							DataRow rowNext = dataTable.Rows[x];
							string treatcodeNext = rowNext["TREATCODE"].ToString();
							if (!treatcodeNext.Equals(treatcode))
								break;

							string kodoperNext = rowNext["KODOPER"].ToString();
							if (kodoperNext.Equals(serviceKodoper)) {
								isDoubled = true;
								rowNext["COMMENT_3"] = "Дубли услуг";
							}
						}

						if (isDoubled) {
							row["COMMENT_3"] = "Дубли услуг";
							continue;
						}
					}

					string mkbCode = row["MKB"].ToString();
					if (!string.IsNullOrEmpty(mkbCode)) {
						string[] mkbCodeSplitted = mkbCode.Split(' ');
						if (mkbUninsured.Contains(mkbCodeSplitted[0])) {
							row["COMMENT_3"] = "Нестраховые заболевания";
							continue;
						}
					}
				} catch (Exception e) {
					Logging.ToLog(e.ToString() + Environment.NewLine + e.StackTrace);
				}
			}
		}
	}
}
