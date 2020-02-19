using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MISReports.ExcelHandlers {
	class Workload : ExcelGeneral {
		public static bool Process(string resultFile) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws, "Услуги Мет. 1"))
				return false;

			try {
				ws.Activate();
				ws.Range["CC2:CE2"].Select();
				xlApp.Selection.AutoFill(ws.Range["CC2:CE" + ws.UsedRange.Rows.Count]);
				ws.Range["CC3:CE3"].Select();
				xlApp.Selection.AutoFill(ws.Range["CC2:CE3"]);
				ws.Range["A2:CB2"].Select();
				xlApp.Selection.Copy();
				ws.Range["A3:CB" + ws.UsedRange.Rows.Count].Select();
				xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
				ws.Range["A1"].Select();

				ws = wb.Sheets["Искл. услуги"];
				ws.Activate();
				ws.Range["A2:K2"].Select();
				xlApp.Selection.Copy();
				ws.Range["A3:K" + ws.UsedRange.Rows.Count].Select();
				xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
				ws.Range["A1"].Select();

				ws = wb.Sheets["Расчет"];
				ws.Activate();
				ws.Range["AB2:AO2"].Select();
				xlApp.Selection.AutoFill(ws.Range["AB2:AO" + ws.UsedRange.Rows.Count]);
				ws.Range["AB3:AO3"].Select();
				xlApp.Selection.AutoFill(ws.Range["AB2:AO3"]);
				ws.Range["A2:AA2"].Select();
				xlApp.Selection.Copy();
				ws.Range["A3:AA" + ws.UsedRange.Rows.Count].Select();
				xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
				ws.Range["Z2:Z2"].Select();
				xlApp.Selection.AutoFill(ws.Range["Z2:Z" + ws.UsedRange.Rows.Count]);

				xlApp.ScreenUpdating = false;

				List<string> deptsToExclude = new List<string> {
					"АНЕСТЕЗИОЛОГИЯ-РЕАНИМАТОЛОГИЯ",
					"ДЕЖУРНЫЙ ВРАЧ",
					"Дежурный врач детский",
					"ПРОЦЕДУРНЫЙ",
					"Процедурный кабинет детский",
					"ФИЗИОПРОЦЕДУРЫ",
					"ТЕЛЕМЕДИЦИНА",
					"Аппаратная коррекция зрения",
					"Процедурный (детский)",
					"Физиопроцедуры детские",
					"ЗДРАВПУНКТ",
					"ОБЩЕЕ",
					"ОМС",
					"ПРЕДРЕЙСОВЫЙ ОСМОТР",
					"СКОРАЯ МЕДИЦИНСКАЯ ПОМОЩЬ"
				};

				List<string> docPostsToExclude = new List<string>() {
					"Администратор",
					"Администратор (старший)",
					"Администратор-регистратор",
					"Администратор-регистратор-кассир",
					"Администрация",
					"Акушерка",
					"Архивариус",
					"Врач (общее)",
					"Главный врач",
					"Заведующий регистратурой",
					"Зам.главного врача по КЭР",
					"Лаборант",
					"Лаборатория",
					"Медицинская сестра",
					"Медицинская сестра (косметология)",
					"Медицинский брат",
					"Сотрудник КДЛ",
					"Сотрудник УК",
					"Фельдшер"
				};

				string filialNameBlock = string.Empty;
				string departmentBlock = string.Empty;
				int firstRowBlock = 0;
				int firstRowFilial = 0;
				int row = 0;

				for (row = 2; row <= ws.UsedRange.Rows.Count + 1; row++) {
					Console.WriteLine("row: " + row + " / " + ws.UsedRange.Rows.Count);

					try {
						string department = ws.Range["F" + row].Value2;
						string filialName = Convert.ToString(ws.Range["D" + row].Value2);

						if (department == null)
							department = string.Empty;

						if (string.IsNullOrEmpty(departmentBlock)) {
							departmentBlock = department;
							filialNameBlock = filialName;
							firstRowBlock = row;
							firstRowFilial = row;
						} 
						
						if (!departmentBlock.Equals(department)) {
							Console.WriteLine("DepartmentTotals: " + departmentBlock);
							CreateDepartmentTotals(wb, ws, xlApp, firstRowBlock, ref row, deptsToExclude.Contains(departmentBlock));
							departmentBlock = department;
							firstRowBlock = row;
						} 
						
						if (!filialNameBlock.Equals(filialName)) {
							Console.WriteLine("FilialTotals: " + filialNameBlock);
							CreateDepartmentTotals(wb, ws, xlApp, firstRowFilial, ref row, false, isMethodic1Total: true);
							CreateDepartmentTotals(wb, ws, xlApp, firstRowFilial, ref row, false, isMethodic2Total: true);
							filialNameBlock = filialName;
							departmentBlock = department;
							firstRowBlock = row;
							firstRowFilial = row;
						}

						if (string.IsNullOrEmpty(department))
							continue;

						if (department.ToLower().Equals("рефлексотерапия")) {
							double filID = ws.Range["C" + row].Value;
							if (filID == 1 || filID == 5 || filID == 12) {
								double chairsCount = 2;
								if (filID == 5) {
									chairsCount = 3;
									ws.Range["Z" + row].Value2 = "4";
									ws.Range["Z" + row].Interior.ColorIndex = 45;
								}

								double timeDS = ws.Range["L" + row].Value2;
								ws.Range["L" + row].FormulaLocal = "=" + timeDS + "/" + chairsCount;
								ws.Range["L" + row].Interior.ColorIndex = 45;

								double timeSchRez = ws.Range["N" + row].Value2;
								ws.Range["N" + row].FormulaLocal = "=" + timeSchRez + "/" + chairsCount;
								ws.Range["N" + row].Interior.ColorIndex = 45;
							}
						}

						if (deptsToExclude.Contains(department)) {
							ws.Range["AN" + row].Value2 = 1;
							continue;
						}

						string docPost = ws.Range["K" + row].Value2;
						if (string.IsNullOrEmpty(docPost))
							continue;

						if (docPostsToExclude.Contains(docPost)) {
							string deptLow = department.ToLower();
							if (!(deptLow.Contains("массаж") ||
								deptLow.Contains("водолечение") ||
								deptLow.Contains("грязелечение") ||
								deptLow.Contains("лечебная физкультура") ||
								deptLow.Contains("медицинская реабилитация"))) {
								ws.Range["AN" + row].Value2 = 1;
								continue;
							}
						}

						if (docPost.Equals("Рентгенолаборант") &&
							!filialName.Equals("17") &&
							!filialName.Equals("15")) {
							ws.Range["AN" + row].Value2 = 1;
							continue;
						}

						if (docPost.Equals("Мануальный терапевт") &&
							((string)ws.Range["H" + row].Value2).StartsWith("Пеньтковский"))
							ws.Range["AN" + row].Value2 = 1;

					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					}
				}

				row--;
				CreateDepartmentTotals(wb, ws, xlApp, 2, ref row, false, isMethodic1Total: true, isGeneralTotal:true);
				CreateDepartmentTotals(wb, ws, xlApp, 2, ref row, false, isMethodic2Total: true, isGeneralTotal:true);

				ws.Columns["AO:AO"].Select();
				xlApp.Selection.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: "=ДЛСТР(СЖПРОБЕЛЫ(AO1))=0");
				xlApp.Selection.FormatConditions(xlApp.Selection.FormatConditions.Count).SetFirstPriority();
				xlApp.Selection.FormatConditions(1).Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
				xlApp.Selection.FormatConditions(1).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
				xlApp.Selection.FormatConditions(1).Interior.TintAndShade = 0;

				ws.Range["A1"].Select();
				xlApp.ActiveWindow.SmallScroll(-10000);

				xlApp.ActiveWindow.ScrollColumn = 8;
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			try {
				WorkloadAddPivotTable(wb, ws, xlApp);
			} catch (Exception e) {
				Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
			}

			xlApp.ScreenUpdating = true;
			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static void CreateDepartmentTotals(Excel.Workbook wb,
											 Excel.Worksheet ws,
											 Excel.Application xlApp,
											 int firstRowBlock,
											 ref int nextBlockFirstRow,
											 bool isNeedToIgnore,
											 bool isMethodic1Total = false,
											 bool isMethodic2Total = false,
											 bool isGeneralTotal = false) {
			Console.WriteLine("Создание итогов по отделению, isMethodic1Total: " + isMethodic1Total + 
				", isMethodic2Total: " + isMethodic2Total + 
				", isGeneralTotal: " + isGeneralTotal);

			//Добавление пустой строки для итогов
			ws.Rows[nextBlockFirstRow + ":" + nextBlockFirstRow].Select();
			xlApp.Selection.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

			//Если итоги для методики №2, то предыдущая строка с данными -2 от текущей
			int prevRow = nextBlockFirstRow - 1;
			if (isMethodic2Total)
				prevRow--;

			//Копирование форматов предыдущей строки
			ws.Range["A" + prevRow + ":AO" + prevRow].Select();
			xlApp.Selection.Copy();
			ws.Range["A" + nextBlockFirstRow + ":AO" + nextBlockFirstRow].PasteSpecial(Excel.XlPasteType.xlPasteFormats);

			//Копирование заголовков предыдущей строки
			ws.Range["A" + firstRowBlock + ":F" + firstRowBlock].Select();
			xlApp.Selection.Copy();
			ws.Range["A" + nextBlockFirstRow + ":F" + nextBlockFirstRow].Select();
			xlApp.ActiveSheet.Paste();

			//Изменение имени отделения
			string departmentTotalName = ws.Range["F" + nextBlockFirstRow].Value2 + " - ИТОГО";
			if (isMethodic1Total)
				departmentTotalName = "_Методика №1 ИТОГО";
			else if (isMethodic2Total)
				departmentTotalName = "_Методика №2 ИТОГО";

			ws.Range["F" + nextBlockFirstRow].Value2 = departmentTotalName;

			//Добавление толстой границы
			AddBoldBorder(ws.Range["A" + firstRowBlock + ":AO" + nextBlockFirstRow]);

			if (isMethodic1Total || isMethodic2Total)
				AddBoldBorder(ws.Range["A" + nextBlockFirstRow + ":AO" + nextBlockFirstRow]);
			
			//Выделение строки итогов цветом
			double tintAndShade = 0.799981688894314;
			if (isMethodic1Total || isMethodic2Total)
				tintAndShade = 0.599993896298105;
			
			if (isGeneralTotal)
				tintAndShade = 0.399975585192419;

			AddInteriorColor(ws.Range["A" + nextBlockFirstRow + ":K" + nextBlockFirstRow], Excel.XlThemeColor.xlThemeColorAccent6, tintAndShade);

			//Формулы для объединения значений базовых данных
			string formulaSumIfLeft = "=SUMIF($AN" + firstRowBlock + ":$AN" + prevRow + ",\"<>1\",L" + firstRowBlock + ":L" + prevRow + ")";
			string formulaSumIfRight = "=SUMIF($AN" + firstRowBlock + ":$AN" + prevRow + ",\"<>1\",AE" + firstRowBlock + ":AE" + prevRow + ")";

			if (isMethodic1Total || isMethodic2Total) {
				formulaSumIfLeft =
					"=SUMIFS(L" + firstRowBlock + ":L" + prevRow +
					",$AN" + firstRowBlock + ":$AN" + prevRow +
					",\"<>1\",$F" + firstRowBlock + ":$F" + prevRow +
					",\"* - ИТОГО\",$AM" + firstRowBlock + ":$AM" + prevRow + ",\"";
				formulaSumIfRight = "=SUMIFS(AE" + firstRowBlock + ":AE" + prevRow +
					",$AN" + firstRowBlock + ":$AN" + prevRow +
					",\"<>1\",$F" + firstRowBlock + ":$F" + prevRow +
					",\"* - ИТОГО\",$AM" + firstRowBlock + ":$AM" + prevRow + ",\"";

				if (isMethodic1Total) {
					formulaSumIfLeft += "=1\"";
					formulaSumIfRight += "=1\"";
				} else if (isMethodic2Total) {
					formulaSumIfLeft += "<>1\"";
					formulaSumIfRight += "<>1\"";
				}

				formulaSumIfLeft += ")";
				formulaSumIfRight += ")";
			}

			//Протягивание формул на соседние ячейки
			ws.Range["L" + nextBlockFirstRow].Formula = formulaSumIfLeft;
			ws.Range["L" + nextBlockFirstRow].Select();
			xlApp.Selection.AutoFill(ws.Range["L" + nextBlockFirstRow + ":Y" + nextBlockFirstRow]);

			ws.Range["AE" + nextBlockFirstRow].Formula = formulaSumIfRight;
			ws.Range["AE" + nextBlockFirstRow].Select();
			xlApp.Selection.AutoFill(ws.Range["AE" + nextBlockFirstRow + ":AK" + nextBlockFirstRow]);

			//План по кол-ву пациентов для отделения для обычных отделений и итогов по методике 2
			if (!isMethodic1Total) {
				ws.Range["AA" + nextBlockFirstRow].Formula = "=(L" + nextBlockFirstRow + "-N" + nextBlockFirstRow + ")*Z" + nextBlockFirstRow;
				ws.Range["AA" + nextBlockFirstRow].AddComment("План по кол-ву пациентов для отделения");
			}

			//Протягивание формул для итогов отделения
			if (!isMethodic1Total && !isMethodic2Total) {
				string[] rangesToFill = new string[] { "Z@:Z$", "AB@:AD$", "AL@:AM$", "AO@:AO$" };
				foreach (string rangeToFill in rangesToFill) {
					string rangeSrc = rangeToFill.Replace("@", prevRow.ToString()).Replace("$", prevRow.ToString());
					string rangeDst = rangeToFill.Replace("@", prevRow.ToString()).Replace("$", nextBlockFirstRow.ToString());

					ws.Range[rangeSrc].Select();
					xlApp.Selection.AutoFill(ws.Range[rangeDst], Excel.XlAutoFillType.xlFillValues);

					//Снятие сообщения об ошибках с ячеек
					foreach (Excel.Range cell in ws.Range["AE" + firstRowBlock + ":AL" + nextBlockFirstRow].Cells)
						if (cell.Errors.Item[Excel.XlErrorChecks.xlInconsistentFormula].Value)
							cell.Errors.Item[Excel.XlErrorChecks.xlInconsistentFormula].Ignore = true;
				}
			}

			//Установка отметки Не учитывать для отделения
			if (isNeedToIgnore)
				ws.Range["AN" + nextBlockFirstRow].Value2 = 1;

			if (isMethodic1Total) {
				//Установка отметки Расчет по методике №1
				ws.Range["AM" + nextBlockFirstRow].Value2 = 1;

				//Формула расчета загрузки для итогов методики №1
				string formulaCount = "=IF(AN" + nextBlockFirstRow + "=1,\"\",IF(AM" + nextBlockFirstRow + 
					"=1,AL" + nextBlockFirstRow +",IFERROR(AI" + nextBlockFirstRow + 
					"*100/((L" + nextBlockFirstRow + "-N" + nextBlockFirstRow + 
					")*Z" + nextBlockFirstRow + "),0)))";
				ws.Range["AO" + nextBlockFirstRow].Formula = formulaCount;

				string formulaCount1 = "=IFERROR(AK" + nextBlockFirstRow + "*100/AJ" + nextBlockFirstRow +",0)";
				ws.Range["AL" + nextBlockFirstRow].Formula = formulaCount1;
			} else if (isMethodic2Total) {
				//Формула расчета плана по кол-ву пациентов для итогов методики №2
				string formulaSumPlan = "=SUMIFS(AA" + firstRowBlock + ":AA" + prevRow +
					",$AN" + firstRowBlock + ":$AN" + prevRow +
					",\"<>1\",$F" + firstRowBlock + ":$F" + prevRow +
					",\"* - ИТОГО\",$AM" + firstRowBlock + ":$AM" + prevRow + ",\"<>1\")";
				ws.Range["AA" + nextBlockFirstRow].Formula = formulaSumPlan;

				//Формула расчета загрузки для итогов методики №2
				string formulaCount = "=IF(AN" + nextBlockFirstRow + "=1,\"\",IF(AM" +
					nextBlockFirstRow + "=1,AL" + nextBlockFirstRow + ",IFERROR(AI" +
					nextBlockFirstRow + "*100/AA" + nextBlockFirstRow + ",0)))";
				ws.Range["AO" + nextBlockFirstRow].Formula = formulaCount;
			}

			if (isMethodic1Total || isMethodic2Total)
				//Снятие уведомления об ошибке в формуле
				if (ws.Range["AO" + nextBlockFirstRow].Errors.Item[Excel.XlErrorChecks.xlInconsistentFormula].Value)
					ws.Range["AO" + nextBlockFirstRow].Errors.Item[Excel.XlErrorChecks.xlInconsistentFormula].Ignore = true;

			//Замена имени отделения для общих итогов
			if (isGeneralTotal) {
				ws.Range["B" + nextBlockFirstRow + ":E" + nextBlockFirstRow].Value2 = string.Empty;
				ws.Range["D" + nextBlockFirstRow].Value2 = "Все клиники";
			}

			nextBlockFirstRow++;
		}

		private static void WorkloadAddPivotTable(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			ws.Cells[1, 1].Select();

			string pivotTableName = @"WorkloadPivotTable";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная таблица"];

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable.PivotFields("ОТДЕЛЕНИЕ").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("ОТДЕЛЕНИЕ").Position = 1;

			pivotTable.PivotFields("ФИЛИАЛ").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
			pivotTable.PivotFields("ФИЛИАЛ").Position = 1;

			pivotTable.AddDataField(pivotTable.PivotFields("Загрузка"),
				"Средняя загрузка", Excel.XlConsolidationFunction.xlAverage);
			pivotTable.PivotFields("Средняя загрузка").NumberFormat = "# ##0,00";

			pivotTable.PivotFields("Не учитывать (=1)").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
			pivotTable.PivotFields("Не учитывать (=1)").Position = 1;
			pivotTable.PivotFields("Не учитывать (=1)").ClearAllFilters();
			pivotTable.PivotFields("Не учитывать (=1)").CurrentPage = "(blank)";

			wsPivote.Activate();
			wsPivote.Columns["B:N"].Select();
			xlApp.Selection.ColumnWidth = 10;

			wsPivote.Range["A1"].Select();

			pivotTable.HasAutoFormat = false;
			pivotTable.ShowTableStyleColumnStripes = true;
			pivotTable.TableStyle2 = "PivotStyleMedium2";
			pivotTable.ColumnGrand = false;
			pivotTable.RowGrand = false;

			foreach (Excel.PivotItem item in pivotTable.PivotFields("ОТДЕЛЕНИЕ").PivotItems())
				if (!item.Name.Contains("ИТОГО"))
					item.Visible = false;

			wsPivote.Columns["B:N"].Select();
			xlApp.Selection.FormatConditions.AddColorScale(3);
			xlApp.Selection.FormatConditions(xlApp.Selection.FormatConditions.Count).SetFirstPriority();

			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(1).Type = Excel.XlConditionValueTypes.xlConditionValueNumber;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(1).Value = 0;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor.TintAndShade = 0;

			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(2).Type = Excel.XlConditionValueTypes.xlConditionValueNumber;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 70;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = 13562593;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor.TintAndShade = 0;

			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(3).Type = Excel.XlConditionValueTypes.xlConditionValueNumber;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 150;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = 6870690;
			xlApp.Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor.TintAndShade = 0;

			wsPivote.Range["C1"].Select();

			wb.ShowPivotTableFieldList = false;
		}
	}
}
