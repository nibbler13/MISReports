using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports.ExcelHandlers {
	class TreatmentsDetails : ExcelGeneral {
		private DataTable dataTable;
		private DataRow dataRow;
		private List<Func<bool>> rules = new List<Func<bool>>();
		private List<string> servicesCodesPregnant = new List<string>();
		private List<string> servicesCodesVaccineGeneral = new List<string>();
		private List<string> servicesCodesVaccineFlu = new List<string>();
		private List<string> servicesCodesVaccineAdult = new List<string>();
		private List<string> servicesCodesVaccineKids = new List<string>();
		private List<string> servicesCodesMaternityInspection = new List<string>();
		private List<string> servicesCodesDroppers = new List<string>();
		private List<string> servicesCodesDoubles = new List<string>();
		private List<string> servicesCodesMRTGeneral = new List<string>();
		private List<string> servicesCodesMRTKids = new List<string>();
		private List<string> servicesCodesKT = new List<string>();
		private List<string> servicesCodesKLKT = new List<string>();
		private List<string> servicesCodesPNDAdult = new List<string>();
		private List<string> servicesCodesPNDKids = new List<string>();
		private List<string> mkbCodesUninsured = new List<string>();
		private int i = 0;
		private int maxKidAge = 0;
		private string fileNameCodes;

		//public enum CompanyName {
		//	АбсолютСтрахование,
		//	Альфастрахование,
		//	Альянс,
		//	БестДоктор,
		//	ВСК,
		//	ВТБ,
		//	ДругиеСК,
		//	ИнгосстрахВзр,
		//	ИнгосстрахДет,
		//	ЛибертиСтрахование,
		//	Метлайф,
		//	Росгосстрах,
		//	Ренессанс,
		//	РЕСО,
		//	СМПСтрахование,
		//	СОГАЗ,
		//	Согласие,
		//	Энергогарант
		//}

		public void PerformDataTable(DataTable dataTable, ReportsInfo.Type type) {
			if (dataTable == null)
				return;

			this.dataTable = dataTable;

			switch (type) {
				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsAbsolut:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleFranchise,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RuleVaccineKids,
						RuleMaterityInspection,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes = 
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\" + 
						@"142 - МЭЭ\Правила\Перечень СК\АбсолютСтрахование №158-М-2016 от 01.12.2016\АбсолютСтрахование.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "ДекретированныеОсмотры", servicesCodesMaternityInspection);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsAlfa:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleKids,
						RuleFranchise,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineGeneral,
						RuleMRTGeneral,
						RuleKT,
						RuleKLKT,
						RulePNDAdult,
						RuleDoubles,
						RuleUninsured
					};

					maxKidAge = 17;

					fileNameCodes =
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\" +
						@"142 - МЭЭ\Правила\Перечень СК\Альфастрахование №492 от 16.05.2005\Альфастрахование.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "Вакцинация", servicesCodesVaccineGeneral);
					ReadWorksheetColumn0(fileNameCodes, "МРТ", servicesCodesMRTGeneral);
					ReadWorksheetColumn0(fileNameCodes, "КТ", servicesCodesKT);
					ReadWorksheetColumn0(fileNameCodes, "КЛКТ", servicesCodesKLKT);
					ReadWorksheetColumn0(fileNameCodes, "ПНД Взрослые", servicesCodesPNDAdult);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsAlliance:
					rules = new List<Func<bool>> { 
						RuleZeroCost,
						RuleGarantyMailAlliance,
						RulePrepaidExpense,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVIP,
						RuleVaccineGeneral,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes = 
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\" +
						@"142 - МЭЭ\Правила\Перечень СК\Альянс №Д-796205-31-03-30 от 20.07.2005\Альянс.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "Вакцинация", servicesCodesVaccineGeneral);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsBestdoctor:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RuleVaccineKids,
						RuleMaterityInspection,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes =
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\" +
						@"142 - МЭЭ\Правила\Перечень СК\БестДоктор №206-77-2017 от 01.02.2018\БестДоктор.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "ДекретированныеОсмотры", servicesCodesMaternityInspection);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsVsk:
					rules = new List<Func<bool>> { 
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RuleVaccineKids,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes = 
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\" + 
						@"142 - МЭЭ\Правила\Перечень СК\ВСК №17000SMM00019 от 01.03.2017\ВСК.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsVtb:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleFranchise,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RuleVaccineKids,
						RuleMaterityInspection,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes = 
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\" + 
						@"142 - МЭЭ\Правила\Перечень СК\ВТБ №77МП16-2908 от 19.12.16\ВТБ.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "ДекретированныеОсмотры", servicesCodesMaternityInspection);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsOther:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleFranchise,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RuleVaccineKids,
						RulePNDKids,
						RulePNDAdult,
						RuleMaterityInspection,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes = 
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\" + 
						@"142 - МЭЭ\Правила\Перечень СК\Другие СК\Другие.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "ПНД дети", servicesCodesPNDKids);
					ReadWorksheetColumn0(fileNameCodes, "ПНД взрослые", servicesCodesPNDAdult);
					ReadWorksheetColumn0(fileNameCodes, "ДекретированныеОсмотры", servicesCodesMaternityInspection);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);
					
					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsIngosstrakhAdult:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleGoldStandard,
						RuleEarlyDiagnostic,
						RuleVIP,
						RuleWorldwideInsurance,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RulePNDAdult,
						RuleMRTGeneral,
						RuleKT,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes =
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\" +
						@"Перечень СК\Ингосстрах взр №6187095-19-18 от 01.08.2018\Ингосстрах взр..xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ПНД взрослые", servicesCodesPNDAdult);
					ReadWorksheetColumn0(fileNameCodes, "МРТ", servicesCodesMRTGeneral);
					ReadWorksheetColumn0(fileNameCodes, "КТ", servicesCodesKT);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsIngosstrakhKid:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleEarlyDiagnostic,
						RuleVIP,
						RuleVaccineFlu,
						RuleVaccineKids,
						RuleMaterityInspection,
						RuleMRTGeneral,
						RuleMRTKids,
						RuleKT,
						RulePNDKids,
						RulePNDAdult,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes = 
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\" + 
						@"Перечень СК\Ингосстрах дет.№6187136-19-18 от 01.08.2018\Ингосстрах дет..xlsx";

					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "ДекретированныеОсмотры", servicesCodesMaternityInspection);
					ReadWorksheetColumn0(fileNameCodes, "МРТ", servicesCodesMRTGeneral);
					ReadWorksheetColumn0(fileNameCodes, "МРТ дети", servicesCodesMRTKids);
					ReadWorksheetColumn0(fileNameCodes, "КТ", servicesCodesKT);
					ReadWorksheetColumn0(fileNameCodes, "ПНД дети", servicesCodesPNDKids);
					ReadWorksheetColumn0(fileNameCodes, "ПНД взрослые", servicesCodesPNDAdult);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);
									   
					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsLiberty:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleFranchise,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RuleVaccineKids,
						RuleMaterityInspection,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes =
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\" +
						@"Перечень СК\ЛибертиСтрахование №0044-17 от 16.05.2017\ЛибертиСтрахование.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "ДекретированныеОсмотры", servicesCodesMaternityInspection);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsMetlife:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleFranchise,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RuleVaccineKids,
						RuleMaterityInspection,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes =
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\" +
						@"Перечень СК\Метлайф №GMD-03164-05-17 от 01.06.2016\Метлайф.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "ДекретированныеОсмотры", servicesCodesMaternityInspection);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsRosgosstrakh:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleFranchise,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RuleVaccineKids,
						RuleMaterityInspection,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes =
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\" +
						@"Перечень СК\Росгосстрах №М-77-Н-ПС-А-2014260 от 21.08.2014г\Росгосстрах.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "ДекретированныеОсмотры", servicesCodesMaternityInspection);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsRenessans:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleFranchise,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineGeneral,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes =
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\" +
						@"Перечень СК\СК Ренессанс №29-17 от 23.05.2017\Ренессанс.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "Вакцинация", servicesCodesVaccineGeneral);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsReso:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleKids,
						RuleFranchise,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineGeneral,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes =
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\" +
						@"Перечень СК\СК РЕСО №17-29 от 01.07.2017\!Правила.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "Вакцинация", servicesCodesVaccineGeneral);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsSmp:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleFranchise,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RuleVaccineKids,
						RuleMaterityInspection,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes =
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\" +
						@"Перечень СК\СМП страхование №4-0019 от 01.03.2017\СМП страхование.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "ДекретированныеОсмотры", servicesCodesMaternityInspection);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsSogaz:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleFranchise,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RuleVaccineKids,
						RuleMaterityInspection,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes =
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\" +
						@"Перечень СК\СОГАЗ №18QP 2124 от 26.02.2019\СОГАЗ.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "ДекретированныеОсмотры", servicesCodesMaternityInspection);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsSoglasie:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RuleVaccineKids,
						RuleMaterityInspection,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes =
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\" +
						@"Перечень СК\Согласие №331610-14314 от 01.06.2017 от 01.07.2017\Согласие.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "ДекретированныеОсмотры", servicesCodesMaternityInspection);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				//------------
				//checked 14.11.2019
				//------------
				case ReportsInfo.Type.TreatmentsDetailsEnergogarant:
					rules = new List<Func<bool>> {
						RuleZeroCost,
						RuleGarantyMail,
						RulePrepaidExpense,
						RuleFranchise,
						RuleVIP,
						RuleProgramForPregnant,
						RuleServicesForPregnant,
						RuleVaccineFlu,
						RuleVaccineAdult,
						RuleVaccineKids,
						RuleMaterityInspection,
						RuleDroppers,
						RuleDoubles,
						RuleUninsured
					};

					fileNameCodes =
						@"\\mskv-fs-01\MSKV Files\Управление информационных технологий\08_Проекты\142 - МЭЭ\Правила\" +
						@"Перечень СК\Энергогарант № М-370 от 15.03.2017\Энергогарант.xlsx";

					ReadWorksheetColumn0(fileNameCodes, "БеременностьУслуги", servicesCodesPregnant);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияГрипп", servicesCodesVaccineFlu);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияВзрослые", servicesCodesVaccineAdult);
					ReadWorksheetColumn0(fileNameCodes, "ВакцинацияНацкалендарь", servicesCodesVaccineKids);
					ReadWorksheetColumn0(fileNameCodes, "ДекретированныеОсмотры", servicesCodesMaternityInspection);
					ReadWorksheetColumn0(fileNameCodes, "Капельницы", servicesCodesDroppers);
					ReadWorksheetColumn0(fileNameCodes, "Дубли", servicesCodesDoubles);
					ReadWorksheetColumn0(fileNameCodes, "Нестраховые", mkbCodesUninsured);

					break;

				default:
					break;
			}

			for (i = 0; i < dataTable.Rows.Count; i++) {
				try {
					dataRow = dataTable.Rows[i];
					string comment_3 = dataRow["COMMENT_3"].ToString();
					if (!string.IsNullOrEmpty(comment_3))
						continue;

					foreach (Func<bool> rule in rules) 
						if (rule()) break;
				} catch (Exception e) {
					Logging.ToLog(e.ToString() + Environment.NewLine + e.StackTrace);
				}
			}
		}

		private void ReadWorksheetColumn0(string fileFullPath, string sheetName, List<string> list) {
			Logging.ToLog("Считывание файла: " + fileFullPath);
			if (File.Exists(fileFullPath)) {
				try {
					DataTable dataTable = ReadExcelFile(fileFullPath, sheetName);
					foreach (DataRow row in dataTable.Rows) {
						string value = row[0].ToString();
						if (string.IsNullOrEmpty(value))
							continue;

						list.Add(value);
					}

					Logging.ToLog("Считано строк: " + list.Count);
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			} else {
				Logging.ToLog("Не удается найти файл (получить доступ): " + fileFullPath);
			}
		}


		private bool RuleZeroCost() {
			string amountrub_a = dataRow["AMOUNTRUB_A"].ToString();
			if (string.IsNullOrEmpty(amountrub_a) || (double.TryParse(amountrub_a, out double serviceCost) && serviceCost == 0)) {
				dataRow["COMMENT_3"] = "Нулевые";
				return true;
			}

			return false;
		}

		private bool RuleGarantyMail() {
			string programType = dataRow["PRG"].ToString().ToLower();
			if (programType.Contains("гарантийное письмо")) {
				dataRow["COMMENT_3"] = "ГП";
				return true;
			}

			return false;
		}
		private bool RuleGarantyMailAlliance() {
			string programType = dataRow["PRG"].ToString().ToLower();
			if (programType.TrimStart(' ').TrimEnd(' ').Equals("гарантийное письмо")) {
				dataRow["COMMENT_3"] = "ГП";
				return true;
			}

			return false;
		}

		private bool RulePrepaidExpense() {
			string programType = dataRow["PRG"].ToString().ToLower();
			if (programType.StartsWith("а") || programType.Contains("аванс")) {
				dataRow["COMMENT_3"] = "Аванс";
				return true;
			}

			return false;
		}

		private bool RuleGoldStandard() {
			string programType = dataRow["PRG"].ToString().ToLower();
			if (programType.Contains("золотой стандарт")) {
				dataRow["COMMENT_3"] = "Золотой стандарт";
				return true;
			}

			return false;
		}

		private bool RuleEarlyDiagnostic() {
			string programType = dataRow["PRG"].ToString().ToLower();
			if (programType.Contains("ранняя диагностика")) {
				dataRow["COMMENT_3"] = "Ранняя диагностика";
				return true;
			}

			return false;
		}

		private bool RuleWorldwideInsurance() {
			string programType = dataRow["PRG"].ToString().ToLower();
			if (programType.Contains("imi")) {
				dataRow["COMMENT_3"] = "Международное страхование";
				return true;
			}

			return false;
		}

		private bool RuleFranchise() {
			string programType = dataRow["PRG"].ToString().ToLower();
			if (programType.Contains("франшиза")) {
				dataRow["COMMENT_3"] = "Франшиза";
				return true;
			}

			return false;
		}

		private bool RuleVIP() {
			string programType = dataRow["PRG"].ToString().ToLower();
			if (programType.Contains("вип") || programType.Contains("vip")) {
				dataRow["COMMENT_3"] = "ВИП";
				return true;
			}

			return false;
		}

		private bool RuleProgramForPregnant() {
			string programType = dataRow["PRG"].ToString().ToLower();
			if (programType.Contains("берем")) {
				dataRow["COMMENT_3"] = "Беременность_программы";
				return true;
			}

			return false;
		}

		private bool RuleMRTGeneral() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesMRTGeneral.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "МРТ";
				return true;
			}

			return false;
		}

		private bool RuleMRTKids() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesMRTKids.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "МРТ дети";
				return true;
			}

			return false;
		}

		private bool RuleKT() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesKT.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "КТ";
				return true;
			}

			return false;
		}

		private bool RuleKLKT() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesKLKT.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "КЛКТ";
				return true;
			}

			return false;
		}

		private bool RulePNDAdult() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesPNDAdult.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "ПНД взрослые";
				return true;
			}

			return false;
		}

		private bool RulePNDKids() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesPNDKids.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "ПНД дети";
				return true;
			}

			return false;
		}

		private bool RuleServicesForPregnant() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesPregnant.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "Беременность_услуги";
				return true;
			}

			return false;
		}

		private bool RuleDroppers() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesDroppers.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "Капельницы";
				return true;
			}

			return false;
		}

		private bool RuleDoubles() {
			string serviceKodoper = dataRow["KODOPER"].ToString();

			if (servicesCodesDoubles.Contains(serviceKodoper)) {
				string treatcode = dataRow["TREATCODE"].ToString();
				bool isDoubled = false;
				for (int x = i + 1; x < dataTable.Rows.Count; x++) {
					DataRow rowNext = dataTable.Rows[x];
					string treatcodeNext = rowNext["TREATCODE"].ToString();
					if (!treatcodeNext.Equals(treatcode))
						break;

					string kodoperNext = rowNext["KODOPER"].ToString();
					if (kodoperNext.Equals(serviceKodoper)) {
						isDoubled = true;
						rowNext["COMMENT_1"] = "Дубли услуг";
					}
				}

				if (isDoubled) {
					dataRow["COMMENT_1"] = "Дубли услуг";
					return true;
				}
			}

			return false;
		}

		private bool RuleUninsured() {
			string mkbCode = dataRow["MKB"].ToString();
			if (!string.IsNullOrEmpty(mkbCode)) {
				string[] mkbCodeSplitted = mkbCode.Split(' ');
				if (mkbCodesUninsured.Contains(mkbCodeSplitted[0])) {
					dataRow["COMMENT_1"] = "Нестраховые заболевания";
					return true;
				}
			}

			return false;
		}

		private bool RuleVaccineGeneral() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesVaccineGeneral.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "Вакцинация";
				return true;
			}

			return false;
		}

		private bool RuleVaccineFlu() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesVaccineFlu.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "Вакцинация грипп";
				return true;
			}

			return false;
		}

		private bool RuleVaccineAdult() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesVaccineAdult.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "Вакцинация взрослые";
				return true;
			}

			return false;
		}

		private bool RuleVaccineKids() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesVaccineKids.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "Вакцинация дети";
				return true;
			}

			return false;
		}

		private bool RuleMaterityInspection() {
			string serviceKodoper = dataRow["KODOPER"].ToString();
			if (servicesCodesMaternityInspection.Contains(serviceKodoper)) {
				dataRow["COMMENT_3"] = "Декретированные осмотры";
				return true;
			}

			return false;
		}

		private bool RuleKids() {
			string age = dataRow["AGE"].ToString();
			if (double.TryParse(age, out double ageParsed) && ageParsed <= maxKidAge) {
				dataRow["COMMENT_3"] = "Дети";
				return true;
			}

			return false;
		}
	}
}
