using MISReports.Properties;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports {
	public class ItemReport {
		public ReportsInfo.Type Type { get; private set; }
		public string Name { get; private set; }
		public string SqlQuery { get; set; }
		public string MailTo { get; private set; }
		public string TemplateFileName { get; private set; }
		public string FolderToSave { get; private set; }
		public string PreviousFile { get; private set; }
		public bool UploadToServer { get; private set; }
		public bool IsSettingsLoaded { get; private set; } = true;
		public string Periodicity { get; private set; }
		public DateTime DateBegin { get; private set; }
		public DateTime DateEnd { get; private set; }
		public string FileResult { get; set; }
		public string FileResultAverageCheckPreviousYear { get; set; }
		public string JIDS { get; private set; }
		public bool UseVerticaDb { get; set; } = false;
		public List<ItemTreatmentsDiscount> TreatmentsDiscounts { get; private set; } = new List<ItemTreatmentsDiscount>();

		public ItemReport(string reportName) {
			Settings settings = Settings.Default;

			if (reportName.Equals(ReportsInfo.Type.FreeCellsDay.ToString())) {
				Type = ReportsInfo.Type.FreeCellsDay;
				SqlQuery = settings.MisDbSqlGetFreeCells;
				MailTo = settings.MailToFreeCellsDay;
				TemplateFileName = settings.TemplateFreeCells;

			} else if (reportName.Equals(ReportsInfo.Type.FreeCellsWeek.ToString())) {
				Type = ReportsInfo.Type.FreeCellsWeek;
				SqlQuery = settings.MisDbSqlGetFreeCells;
				MailTo = settings.MailToFreeCellsWeek;
				TemplateFileName = settings.TemplateFreeCells;

			} else if (reportName.Equals(ReportsInfo.Type.UnclosedProtocolsWeek.ToString())) {
				Type = ReportsInfo.Type.UnclosedProtocolsWeek;
				SqlQuery = settings.MisDbSqlGetUnclosedProtocols;
				MailTo = settings.MailToUnclosedProtocolsMonth;
				TemplateFileName = settings.TemplateUnclosedProtocols;
				FolderToSave = settings.FolderToSaveUnclosedProtocols;

			} else if (reportName.Equals(ReportsInfo.Type.UnclosedProtocolsMonth.ToString())) {
				Type = ReportsInfo.Type.UnclosedProtocolsMonth;
				SqlQuery = settings.MisDbSqlGetUnclosedProtocols;
				MailTo = settings.MailToUnclosedProtocolsMonth;
				TemplateFileName = settings.TemplateUnclosedProtocols;
				FolderToSave = settings.FolderToSaveUnclosedProtocols;

			} else if (reportName.Equals(ReportsInfo.Type.MESUsage.ToString())) {
				Type = ReportsInfo.Type.MESUsage;
				SqlQuery = settings.MisDbSqlGetMESUsage;
				MailTo = settings.MailToMESUsage;
				TemplateFileName = settings.TemplateMESUsage;
				FolderToSave = settings.FolderToSaveMESUsage;

			} else if (reportName.Equals(ReportsInfo.Type.MESUsageFull.ToString())) {
				Type = ReportsInfo.Type.MESUsageFull;
				SqlQuery = settings.MisDbSqlGetMESUsage;
				MailTo = settings.MailToMESUsage;
				TemplateFileName = settings.TemplateMESUsageFull;

			} else if (reportName.Equals(ReportsInfo.Type.OnlineAccountsUsage.ToString())) {
				Type = ReportsInfo.Type.OnlineAccountsUsage;
				SqlQuery = settings.MisDbSqlGetOnlineAccountsUsage;
				MailTo = settings.MailToOnlineAccountsUsage;
				TemplateFileName = settings.TemplateOnlineAccountsUsage;

			} else if (reportName.Equals(ReportsInfo.Type.TelemedicineOnlyIngosstrakh.ToString())) {
				Type = ReportsInfo.Type.TelemedicineOnlyIngosstrakh;
				SqlQuery = settings.MisDbSqlGetTelemedicine;
				TemplateFileName = settings.TemplateTelemedicine;
				MailTo = settings.MailToTelemedicineOnlyIngosstrakh;

			} else if (reportName.Equals(ReportsInfo.Type.TelemedicineAll.ToString())) {
				Type = ReportsInfo.Type.TelemedicineAll;
				SqlQuery = settings.MisDbSqlGetTelemedicine;
				TemplateFileName = settings.TemplateTelemedicine;
				MailTo = settings.MailToTelemedicineAll;

			} else if (reportName.Equals(ReportsInfo.Type.NonAppearance.ToString())) {
				Type = ReportsInfo.Type.NonAppearance;
				SqlQuery = settings.MisDbSqlGetNonAppearance;
				TemplateFileName = settings.TemplateNonAppearance;
				MailTo = settings.MailToNonAppearance;
				FolderToSave = settings.FolderToSaveNonAppearance;

			} else if (reportName.Equals(ReportsInfo.Type.VIP_MSSU.ToString())) {
				Type = ReportsInfo.Type.VIP_MSSU;
				SqlQuery = settings.MisDbSqlGetVIP.Replace("@filialList", "12");
				TemplateFileName = settings.TemplateVIP;
				MailTo = settings.MailToVIP_MSSU;
				PreviousFile = settings.PreviousFileVIP_MSSU;

			} else if (reportName.Equals(ReportsInfo.Type.VIP_Moscow.ToString())) {
				Type = ReportsInfo.Type.VIP_Moscow;
				SqlQuery = settings.MisDbSqlGetVIP.Replace("@filialList", "1,5,12,6");
				TemplateFileName = settings.TemplateVIP;
				MailTo = settings.MailToVIP_Moscow;
				PreviousFile = settings.PreviousFileVIP_Moscow;

			} else if (reportName.Equals(ReportsInfo.Type.VIP_MSKM.ToString())) {
				Type = ReportsInfo.Type.VIP_MSKM;
				SqlQuery = settings.MisDbSqlGetVIP.Replace("@filialList", "1");
				TemplateFileName = settings.TemplateVIP;
				MailTo = settings.MailToVIP_MSKM;
				PreviousFile = settings.PreviousFileVIP_MSKM;

			} else if (reportName.Equals(ReportsInfo.Type.VIP_PND.ToString())) {
				Type = ReportsInfo.Type.VIP_PND;
				SqlQuery = settings.MisDbSqlGetVIP.Replace("@filialList", "6");
				TemplateFileName = settings.TemplateVIP;
				MailTo = settings.MailToVIP_PND;
				PreviousFile = settings.PreviousFileVIP_PND;

			} else if (reportName.Equals(ReportsInfo.Type.RegistryMarks.ToString())) {
				Type = ReportsInfo.Type.RegistryMarks;
				SqlQuery = settings.MisDbSqlGetRegistryMarks;
				TemplateFileName = settings.TemplateRegistryMarks;
				MailTo = settings.MailToRegistryMarks;

			} else if (reportName.Equals(ReportsInfo.Type.Workload.ToString())) {
				Type = ReportsInfo.Type.Workload;
				TemplateFileName = settings.TemplateWorkload;
				MailTo = settings.MailToWorkload;
				FolderToSave = settings.FolderToSaveWorkload;

			} else if (reportName.Equals(ReportsInfo.Type.Robocalls.ToString())) {
				Type = ReportsInfo.Type.Robocalls;
				SqlQuery = settings.MisDbSqlGetRobocalls;
				TemplateFileName = settings.TemplateRobocalls;
				MailTo = settings.MailToRobocalls;

			} else if (reportName.Equals(ReportsInfo.Type.UniqueServices.ToString())) {
				Type = ReportsInfo.Type.UniqueServices;
				SqlQuery = settings.MisDbSqlGetUniqueServices;
				TemplateFileName = settings.TemplateUniqueServices;
				MailTo = settings.MailToUniqueServices;

			} else if (reportName.Equals(ReportsInfo.Type.UniqueServicesRegions.ToString())) {
				Type = ReportsInfo.Type.UniqueServicesRegions;
				SqlQuery = settings.MisDbSqlGetUniqueServicesRegions;
				TemplateFileName = settings.TemplateUniqueServicesRegions;
				MailTo = settings.MailToUniqueServicesRegions;

			} else if (reportName.Equals(ReportsInfo.Type.PriceListToSite.ToString())) {
				Type = ReportsInfo.Type.PriceListToSite;
				SqlQuery = settings.MisDbSqlGetPriceListToSite;
				TemplateFileName = settings.TemplatePriceListToSite;
				MailTo = settings.MailToPriceListToSite;
				FolderToSave = settings.FolderToSavePriceListToSite;
				UploadToServer = true;

			} else if (reportName.Equals(ReportsInfo.Type.GBooking.ToString())) {
				Type = ReportsInfo.Type.GBooking;
				SqlQuery = settings.MisDbSqlGetGBooking;
				TemplateFileName = settings.TemplateGBooking;
				MailTo = settings.MailToGBooking;

			} else if (reportName.Equals(ReportsInfo.Type.PersonalAccountSchedule.ToString())) {
				Type = ReportsInfo.Type.PersonalAccountSchedule;
				SqlQuery = settings.MisDbSqlGetPersonalAccountSchedule;
				TemplateFileName = settings.TemplatePersonalAccountSchedule;
				MailTo = settings.MailToPersonalAccountSchedule;

			} else if (reportName.Equals(ReportsInfo.Type.ProtocolViewCDBSyncEvent.ToString())) {
				Type = ReportsInfo.Type.ProtocolViewCDBSyncEvent;
				SqlQuery = settings.MisDbSqlGetProtocolViewCDBSyncEvent;
				TemplateFileName = settings.TemplateProtocolViewCDBSyncEvent;
				MailTo = settings.MailToProtocolViewCDBSyncEvent;
				FolderToSave = settings.FolderToSaveProtocolViewCDBSyncEvent;

			} else if (reportName.Equals(ReportsInfo.Type.FssInfo.ToString())) {
				Type = ReportsInfo.Type.FssInfo;
				SqlQuery = settings.MisDbSqlGetFssInfo;
				TemplateFileName = settings.TemplateFssInfo;
				MailTo = settings.MailToFssInfo;

			} else if (reportName.Equals(ReportsInfo.Type.TimetableToProdoctorovRu.ToString())) {
				Type = ReportsInfo.Type.TimetableToProdoctorovRu;
				SqlQuery = settings.MisDbSqlGetTimetableToProdoctorovRu;
				MailTo = settings.MailToTimetableToProdoctorovRu;
				UploadToServer = true;

			} else if (reportName.Equals(ReportsInfo.Type.RecordsFromInsuranceCompanies.ToString())) {
				Type = ReportsInfo.Type.RecordsFromInsuranceCompanies;
				SqlQuery = settings.MisDbSqlGetRecordsFromInsuranceCompanies;
				TemplateFileName = settings.TemplateRecordsFromInsuranceCompanies;
				MailTo = settings.MailToRecordsFromInsuranceCompanies;

			} else if (reportName.Equals(ReportsInfo.Type.AverageCheckRegular.ToString())) {
				Type = ReportsInfo.Type.AverageCheckRegular;
				SqlQuery = settings.MisDbSqlGetAverageCheck;
				TemplateFileName = settings.TemplateAverageCheck;
				MailTo = settings.MailToAverageCheck;

			} else if (reportName.Equals(ReportsInfo.Type.AverageCheckIGS.ToString())) {
				Type = ReportsInfo.Type.AverageCheckIGS;
				SqlQuery = settings.MisDbSqlGetAverageCheck;
				TemplateFileName = settings.TemplateAverageCheck;
				MailTo = settings.MailToAverageCheckIGS;

			} else if (reportName.Equals(ReportsInfo.Type.AverageCheckMSK.ToString())) {
				Type = ReportsInfo.Type.AverageCheckMSK;
				SqlQuery = settings.MisDbSqlGetAverageCheckMSK;
				TemplateFileName = settings.TemplateAverageCheck;
				MailTo = settings.MailToAverageCheckMSK;

			} else if (reportName.Equals(ReportsInfo.Type.AverageCheckCash.ToString())) {
				Type = ReportsInfo.Type.AverageCheckCash;
				SqlQuery = settings.MisDbSqlGetAverageCheckCash;
				TemplateFileName = settings.TemplateAverageCheckCash;
				MailTo = settings.MailToAverageCheckCash;

			} else if (reportName.Equals(ReportsInfo.Type.CompetitiveGroups.ToString())) {
				Type = ReportsInfo.Type.CompetitiveGroups;
				SqlQuery = settings.MisDbSqlGetCompetitiveGroups;
				TemplateFileName = settings.TemplateCompetitiveGroups;
				MailTo = settings.MailToCompetitiveGroups;

			} else if (reportName.Equals(ReportsInfo.Type.LicenseStatistics.ToString())) {
				Type = ReportsInfo.Type.LicenseStatistics;
				SqlQuery = settings.MisDbSqlGetLicenseStatistics;
				TemplateFileName = settings.TemplageLicenseStatistics;
				MailTo = settings.MailToLicenseStatistics;
				FolderToSave = settings.FolderToSaveLicenseStatistics;

				//-----------------------------------------------------------------------------------------------------
				#region TreatmentsDetails
			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAbsolut.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAbsolut;
				JIDS = "991515382,991519409,991519865,991523030";
				//991515382   990000581   ООО "Абсолют Страхование"(Москва) Страховая компания  158 - М / 2016								ООО "Абсолют Страхование" - факт / аванс - Москва
				//991519409   990000581   ООО "Абсолют Страхование"(Москва) Страховая компания  ДС № 8 к Дог. № 158 - М / 2016(ВАКЦИНАЦИЯ)	ООО "Абсолют Страхование" - факт - Москва
				//991519865   990000581   ООО "Абсолют Страхование"(Москва) Страховая компания  ДС №10 к Дог.№158 - М / 2016(ВАКЦИНАЦИЯ)	ООО "Абсолют Страхование" - факт - Ступино
				//991523030   990000581   ООО "Абсолют Страхование"(Москва) Страховая компания  ДС №19 к Дог.№158 - М / 2016(ВАКЦИНАЦИЯ)	ООО "Абсолют Страхование" - факт / аванс - Москва

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 5);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

				ItemTreatmentsDiscount discount3 = new ItemTreatmentsDiscount(new DateTime(2021, 1, 1), new DateTime(2021, 6, 30), -1);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount3.AddSmpDeptToExclude();
				discount3.AddDocOnlineTelemedCovidKodoperToExclude();
				discount3.AddCovidInfoToExclude();
				TreatmentsDiscounts.Add(discount3);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAlfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAlfa;
				JIDS = "100005,991520911,991514852";
				//100005		5   АО "АльфаСтрахование"(Москва)  Страховая компания  492									АО "АльфаСтрахование" - факт / аванс - Москва
				//991514852		5   АО "АльфаСтрахование"(Москва)  Страховая компания  ДС №31 / 2016 к Дог.№492(ВАКЦИНАЦИЯ) На оказание медицинских услуг
				//991520911		5   АО "АльфаСтрахование"(Москва)  Страховая компания  492(ВИП ОТДЕЛЕНИЕ)					АО "АльфаСтрахование" - факт - Москва

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2018, 6, 1), null, 20);
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				discount.AddCovidInfoToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 6, 1), new DateTime(2020, 6, 30), 2.9f);
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAlfaSpb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAlfaSpb;
				JIDS = "990424275"; //80/10-09
				//990424275   990000485   АО АльфаСтрахование(Спб)    Страховая компания  80 / 10 - 09    АО "АльфаСтрахование" - факт - Спб

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAlliance.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAlliance;
				JIDS = "991511535,991520499,991519440,991521374,991511568,991511568";
				//991511535   990000332   ООО СК "Альянс Жизнь"(Москва)  Страховая компания  Д - 796205 / 31 - 03 - 30								ООО СК "Альянс Жизнь" - факт / аванс - Москва
				//991511568   990000332   ООО СК "Альянс Жизнь"(Москва)  Страховая компания  ДС №87 к дог.№Д - 796205 / 31 - 03 - 30(ДЕТИ)			ООО СК "Альянс Жизнь" - факт - Москва
				//991520499   990000332   ООО СК "Альянс Жизнь"(Москва)  Страховая компания  Д - 796205 / 31 - 03 - 30(ВИП ОТДЕЛЕНИЕ)				ООО СК "Альянс Жизнь" - факт - Москва
				//991521374   990000332   ООО СК "Альянс Жизнь"(Москва)  Страховая компания  ДС №140 к Дог №Д - 796205 / 31 - 03 - 30(Сheck - up)	ООО СК "Альянс Жизнь" - факт - Москва
				//991519440   990000332   ООО СК "Альянс Жизнь"(Москва)  Страховая компания  ДС №127 к дог. Д - 796205 / 31 - 03 - 30(ВАКЦИНАЦИЯ)   ООО СК "Альянс Жизнь" - факт - Москва

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2021, 12, 31), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				discount.AddCovidInfoToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 2, 1), new DateTime(2020, 3, 31), 20);
				discount2.AddKtMrtPndSmpDeptToExclude();
				TreatmentsDiscounts.Add(discount2);

				ItemTreatmentsDiscount discount3 = new ItemTreatmentsDiscount(new DateTime(2020, 4, 1), new DateTime(2020, 12, 31), 10);
				discount3.AddKtMrtPndSmpDeptToExclude();
				discount3.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount3);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsBestdoctor.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsBestdoctor;
				JIDS = "991520964, 991526106";
				//991520964   990002501   ООО "Бестдоктор"(Москва)   Страховая компания  206 - 77 - 2017									ООО "Бестдоктор" - факт - Москва
				//991526106   990002501   ООО "Бестдоктор"(Москва)   Страховая компания  ДС №12 к дог.№ 206 - 77 - 2017(Вакцинация 2019)    ООО "Бестдоктор" - факт - Москва

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(300000, 700000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(700001, 1500000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1500001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 5);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

				ItemTreatmentsDiscount discount3 = new ItemTreatmentsDiscount(new DateTime(2021, 1, 1), new DateTime(2021, 12, 31), -1);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(300000, 700000), 5);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(700001, 1500000), 10);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(1500001, 3000000), 15);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(3000001, 5000000), 20);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(5000001, 10000000), 25);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(10000001, -1), 27);
				discount3.AddSmpDeptToExclude();
				discount3.AddDocOnlineTelemedCovidKodoperToExclude();
				discount3.AddCovidInfoToExclude();
				discount3.ExcludeKodopers.Add("101944");
				TreatmentsDiscounts.Add(discount3);

				ItemTreatmentsDiscount discount4 = new ItemTreatmentsDiscount(new DateTime(2020, 12, 10), new DateTime(2021, 6, 30), 20, true);
				discount4.ServiceListToApply.Add("101944");
				TreatmentsDiscounts.Add(discount4);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsEnergogarant.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsEnergogarant;
				JIDS = "991523453,991517214,991520348";
				//991517214   990000008   ПАО "САК "ЭНЕРГОГАРАНТ" (Москва)	Страховая компания	М-370								ПАО "САК "ЭНЕРГОГАРАНТ" - факт / аванс - Москва
				//991520348   990000008   ПАО "САК "ЭНЕРГОГАРАНТ" (Москва)	Страховая компания	М-370 (ВИП ОТДЕЛЕНИЕ)				ПАО "САК "ЭНЕРГОГАРАНТ" - факт - Москва
				//991523453   990000008   ПАО "САК "ЭНЕРГОГАРАНТ" (Москва)	Страховая компания	ДС №14 к Дог.№М-370 (ВАКЦИНАЦИЯ)	ПАО "САК "ЭНЕРГОГАРАНТ" - факт - Москва

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhAdult.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhAdult;
				JIDS = "991522348,991522924,991525955,991522442";
				//991522348   4   СПАО "ИНГОССТРАХ"(Москва)  Страховая компания  6187095 - 19 / 18										СПАО "Ингосстрах" - факт - Москва
				//991522442   4   СПАО "ИНГОССТРАХ"(Москва)  Страховая компания  ДС №5 к Дог.№6187095 - 19 / 18(Проверь себя)			СПАО "Ингосстрах" - факт - Москва
				//991522924   4   СПАО "ИНГОССТРАХ"(Москва)  Страховая компания  ДС №10 к дог. 6187095 - 19 / 18(Вакцинация грипп 2018) СПАО "Ингосстрах" - факт - Москва
				//991525955   4   СПАО "ИНГОССТРАХ"(Москва)  Страховая компания  ДС №23 к Дог.6187095 - 19 / 18(Вакцинация 2019)		СПАО "Ингосстрах" - факт - Москва

				//=====================
				//Управленческая скидка
				//=====================
				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 20);
				discount.ExcludeDepartments.Add("СКОРАЯ МЕДИЦИНСКАЯ ПОМОЩЬ");
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 10);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

				ItemTreatmentsDiscount discount3 = new ItemTreatmentsDiscount(new DateTime(2021, 1, 1), new DateTime(2021, 12, 31), -1);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(100000000, 120000000), 5);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(120000001, 150000000), 10);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(150000001, 200000000), 15);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(200000001, -1), 20);
				TreatmentsDiscounts.Add(discount3);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhKid.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhKid;
				JIDS = "991522386";
				//991522386   4   СПАО "ИНГОССТРАХ"(Москва)  Страховая компания  6187136 - 19 / 18   СПАО "Ингосстрах" - факт - Москва

				//=====================
				//Управленческая скидка
				//=====================
				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 20);
				discount.ExcludeDepartments.Add("СКОРАЯ МЕДИЦИНСКАЯ ПОМОЩЬ");
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 10);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

				ItemTreatmentsDiscount discount3 = new ItemTreatmentsDiscount(new DateTime(2021, 1, 1), new DateTime(2021, 12, 31), -1);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(35000001, 45000000), 5);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(45000001, 55000000), 10);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(55000001, 65000000), 15);
				discount3.DynamicDiscount.Add(new Tuple<int, int>(65000001, -1), 20);
				TreatmentsDiscounts.Add(discount3);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsLiberty.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsLiberty;
				JIDS = "991517912";
				//991517912   990000512   АО Совкомбанк страхование(Москва)(Ранее Либерти страхование)  Страховая компания  0044 / 17 АО "Совкомбанк страхование" - факт - Москва

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2021, 1, 1), new DateTime(2021, 6, 30), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000000, 2000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, 4000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(4000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				discount.AddCovidInfoToExclude();
				TreatmentsDiscounts.Add(discount);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsMetlife.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsMetlife;
				JIDS = "991517927,991523451,991519436";
				//991517927   13  АО "МетЛайф"(Москва)   Страховая компания  GMD - 03164 / 05 - 17								АО "МетЛайф" - факт / аванс - Москва
				//991519436   13  АО "МетЛайф"(Москва)   Страховая компания  ДС №7 к Дог.№GMD - 03164 / 05 - 17(ВАКЦИНАЦИЯ)		АО "МетЛайф" - факт - Москва
				//991523451   13  АО "МетЛайф"(Москва)   Страховая компания  ДС №17 к Дог.№GMD - 03164 / 05 - 17(Вакцинация)	АО "МетЛайф" - факт - Москва

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2021, 6, 30), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				discount.AddCovidInfoToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 5);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsOther.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsOther;
				JIDS = "991515382,991519409,991519865,991523030,100005," +
					"991520911,991514852.990424275,991511535,991520499," +
					"991519440,991521374,991511568.991520964,991526106," +
					"991523453,991517214,991520348,991522348,991522924," +
					"991525955,991522442,991522386,991517912,991517927," +
					"991523451,991519436,991523042,991523280,991523170," +
					"991518370,991521272,991523038,991526075,991519595," +
					"991511705,1990097479,991516698,991521960,991524638," +
					"991520913,991518470,991519761,991523028,991516556," +
					"991520387,991523215,991519361,991525970,991515797," +
					"991520427,991512906 ,991357338,991370062,990389345," +
					"991379370,991523486,991523489,991524671,991524697," +
					"991527569,991520964,991511568," + // ТОП-17 СК
					"10021349,991521572,100006," + //ЛМС_0 + ЛМС_6
					"991511059,991511056,991511054,991511055,991511052,991511057"; //ОМС К-Урал

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsRenessans.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsRenessans;
				JIDS = "991523042,991523280,991523170";
				//991523042   990032407   АО "Группа Ренессанс Страхование"(Москва)  Страховая компания  29 / 17 от 23.05.2017						АО "Группа Ренессанс Страхование" - факт / аванс - Москва
				//991523170   990032407   АО "Группа Ренессанс Страхование"(Москва)  Страховая компания  ДС №13 / 14 к дог. № 29 / 17(Chekc - Up)	АО "Группа Ренессанс Страхование" - факт - Москва
				//991523280   990032407   АО "Группа Ренессанс Страхование"(Москва)  Страховая компания  ДС № 20 к дог 29 / 17(ВАКЦИНАЦИЯ)			АО "Группа Ренессанс Страхование" - факт - Москва

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2021, 6, 30), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				discount.AddCovidInfoToExclude();
				discount.ExcludeKodopers.Add("101944");
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 10);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

				ItemTreatmentsDiscount discount3 = new ItemTreatmentsDiscount(new DateTime(2021, 1, 1), new DateTime(2021, 6, 30), 20, true);
				discount3.ServiceListToApply.Add("101944");
				TreatmentsDiscounts.Add(discount3);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsReso.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsReso;
				JIDS = "991518370,991521272,991523038,991526075,991519595";
				//991518370   12  САО "РЕСО-Гарантия"(Москва)    Страховая компания  17 / 29									САО "РЕСО-Гарантия" - факт - Москва
				//991521272   12  САО "РЕСО-Гарантия"(Москва)    Страховая компания  17 / 29(ВИП ОТДЕЛЕНИЕ)						САО "РЕСО-Гарантия" - факт - Москва
				//991519595   12  САО "РЕСО-Гарантия"(Москва)    Страховая компания  ДС №3 к дог. № 17 / 29(ВАКЦИНАЦИЯ)			СПАО "РЕСО-Гарантия" - факт - Москва
				//991523038   12  САО "РЕСО-Гарантия"(Москва)    Страховая компания  ДС №15 к Дог.№17 / 29(ВАКЦИНАЦИЯ)			СПАО "РЕСО-Гарантия" - факт - Москва
				//991526075   12  САО "РЕСО-Гарантия"(Москва)    Страховая компания  ДС №21 к дог.№17 / 29(Вакцинация_2019)		СПАО "РЕСО-Гарантия" - факт - Москва

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2021, 6, 30), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				discount.AddCovidInfoToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 10);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsRosgosstrakh.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsRosgosstrakh;
				JIDS = "991511705,1990097479";
				//1990097479  19			ООО "Росгосстрах"(Москва)		Страховая компания  М - 77 - Н - ПС - А - 2014 / 260_ст				На оказание медицинских услуг
				//991511705   990001956		ПАО СК "Росгосстрах"(Москва)	Страховая компания  М - 77 - Н - ПС - А - 2014 / 260 от 21.08.2014  ПАО СК "Росгосстрах" - факт - Москва

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2021, 6, 30), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				discount.AddCovidInfoToExclude();
				discount.ExcludeKodopers.Add("101944");
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 10);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

				ItemTreatmentsDiscount discount3 = new ItemTreatmentsDiscount(new DateTime(2021, 1, 1), new DateTime(2021, 6, 30), 20, true);
				discount3.ServiceListToApply.Add("101944");
				TreatmentsDiscounts.Add(discount3);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSmp.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSmp;
				JIDS = "991516698,991521960";
				//991516698   990002381   ООО «СМП - Страхование» (Москва)Страховая компания  4 - 0019										ООО «СМП - Страхование»-факт - аванс - Москва
				//991521960   990002381   ООО «СМП - Страхование» (Москва)Страховая компания ДС №12,11 к дог к Дог № 4 - 0019(Chekc - Up)   ООО «СМП - Страхование»-факт - Москва

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSogaz.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSogaz;
				JIDS = "991524638";
				//991524638   21  АО "СОГАЗ"(Москва)    Страховая компания  18 QP 2124 от 26.02.19  АО "СОГАЗ" - факт - Москва

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2021, 6, 30), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				discount.AddCovidInfoToExclude();
				discount.ExcludeKodopers.Add("101944");
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 2, 1), new DateTime(2020, 12, 31), 10);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

				ItemTreatmentsDiscount discount3 = new ItemTreatmentsDiscount(new DateTime(2021, 1, 1), new DateTime(2021, 6, 30), 20, true);
				discount3.ServiceListToApply.Add("101944");
				TreatmentsDiscounts.Add(discount3);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSoglasie.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSoglasie;
				JIDS = "991520913,991518470,991519761,991523028";
				//991520913   138 ООО "СК Согласие"(Москва)  Страховая компания  331610 - 14314(ВИП ОТДЕЛЕНИЕ)    ООО "СК Согласие" - факт - Москва
				//991518470   138 ООО "СК Согласие"(Москва)  Страховая компания  331610 - 14314 от 01.06.2017  ООО "СК Согласие" - факт / аванс - Москва
				//991519761   138 ООО "СК Согласие"(Москва)  Страховая компания  ДС №7 к Дог.№331610 - 14314(ВАКЦИНАЦИЯ)  ООО "СК Согласие" - факт - Москва
				//991523028   138 ООО "СК Согласие"(Москва)  Страховая компания  ДС№21 к дог.№331610 - 14314(ВАКЦИНАЦИЯ)  ООО "СК Согласие" - факт - Москва

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2021, 6, 30), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				discount.AddCovidInfoToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 5);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsVsk.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsVsk;
				JIDS = "991516556,991520387,991523215,991519361,991525970";
				//991516556   107 САО "ВСК"(Москва)  Страховая компания  17000SMM00019									САО "ВСК" - факт / аванс - Москва
				//991520387   107 САО "ВСК"(Москва)  Страховая компания  17000SMM00019(ВИП ОТДЕЛЕНИЕ)					САО "ВСК" - факт - Москва
				//991519361   107 САО "ВСК"(Москва)  Страховая компания  ДС №12 к Дог.№17000SMM00019(ВАКЦИНАЦИЯ)		САО "ВСК" - факт - Москва
				//991523215   107 САО "ВСК"(Москва)  Страховая компания  ДС № 26 к дог.17000SMM00019(ВАКЦИНАЦИЯ)		САО "ВСК" - факт - Москва
				//991525970   107 САО "ВСК"(Москва)  Страховая компания  ДС №№ 41 к дог 17000SMM00019(Вакцинация 2019)  САО "ВСК" - факт - Москва

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2021, 6, 30), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(4000000, 5000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(5000001, 10000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(10000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				discount.AddCovidInfoToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 5);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsVtb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsVtb;
				JIDS = "991515797,991520427";
				//991515797   91  ООО СК "ВТБ-Страхование"(Москва)   Страховая компания  77МП16 - 2908						ООО СК "ВТБ-Страхование" - факт / аванс - Москва
				//991520427   91  ООО СК "ВТБ-Страхование"(Москва)   Страховая компания  77МП16 - 2908(ВИП ОТДЕЛЕНИЕ)		ООО СК "ВТБ-Страхование" - факт - Москва

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAll.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAll;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhSochi.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhSochi;
				JIDS = "991512906"; 
				//991512906	990002076	СПАО "ИНГОССТРАХ" (Сочи)	Страховая компания	4986881-19/16	СПАО "ИНГОССТРАХ"-факт-Сочи

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhKrasnodar.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhKrasnodar;
				JIDS = "991357338"; 
				//991357338   990000625   СПАО "ИНГОССТРАХ"(Краснодар)   Страховая компания	№ 567751 - 19 / 11  СПАО "ИНГОССТРАХ" факт - Краснодар

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhUfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhUfa;
				JIDS = "991370062"; 
				//991370062   990000664   СПАО "ИНГОССТРАХ"(Уфа) Страховая компания	№ 681187 - 19 / 11  СПАО "ИНГОССТРАХ" - факт - Уфа

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhSpb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhSpb;
				JIDS = "990389345";
				//990389345   990000416   СПАО "Ингосстрах"(Спб) Страховая компания  267673 - 19 / 09    СПАО "Ингосстрах" - факт - Спб

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhKazan.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhKazan;
				JIDS = "991379370";
				//991379370   990000708   СПАО «ИНГОССТРАХ» (Казань)Страховая компания  № 714760 - 19 / 11  СПАО «ИНГОССТРАХ» -факт - Казань

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsBestDoctorSpb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsBestDoctorSpb;
				JIDS = "991523486";
				//991523486   990035160   ООО «Бестдоктор» (СПБ)Юридические лица    522 - 78 - 2018 ООО «Бестдоктор»-факт - СПБ

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsBestDoctorUfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsBestDoctorUfa;
				JIDS = "991523489"; //535-02-18
				//991523489   990035162   ООО «Бестдоктор» (Уфа)Юридические лица    535 - 02 - 18   ООО «Бестдоктор»-факт - Уфа

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSogazUfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSogazUfa;
				JIDS = "991524671,991524697";
				//991524671   990001465   АО "СОГАЗ"(Уфа)    Страховая компания  2719RP055													АО "СОГАЗ" - факт - Уфа
				//991524697   990001465   АО "СОГАЗ"(Уфа)    Страховая компания  ДС № 2719RP055 - 02  к дог. №2719RP055(ГК «БАШНЕФТЬ»)		АО "СОГАЗ" - факт - Уфа

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSogazMed.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSogazMed;
				JIDS = "991527569";
				//991527569   990001868   АО СК «СОГАЗ - МЕД» (Москва)Страховая компания  20 QP 268 / SM от 04.09.2020г.АО СК «СОГАЗ - МЕД»-факт - Москва

				#endregion TreatmentsDetails
				//-----------------------------------------------------------------------------------------------------

			} else if (reportName.Equals(ReportsInfo.Type.TimetableToSite.ToString())) {
				Type = ReportsInfo.Type.TimetableToSite;
				SqlQuery = settings.MisDbSqlGetTimetableToSite;
				MailTo = settings.MailToTimetableToSite;
				UploadToServer = true;

			} else if (reportName.Equals(ReportsInfo.Type.MicroSipContactsBook.ToString())) {
				Type = ReportsInfo.Type.MicroSipContactsBook;
				MailTo = settings.MailToMicroSipContactsBook;
				FolderToSave = settings.FolderToSaveMicroSipContactsBook;

			} else if (reportName.Equals(ReportsInfo.Type.TasksForItilium.ToString())) {
				Type = ReportsInfo.Type.TasksForItilium;
				MailTo = settings.MailToTasksForItilium;

			} else if (reportName.Equals(ReportsInfo.Type.FirstTimeVisitPatients.ToString())) {
				Type = ReportsInfo.Type.FirstTimeVisitPatients;
				SqlQuery = settings.MisDbSqlGetFirstTimeVisitPatients;
				MailTo = settings.MailToFirstTimeVisitPatients;
				TemplateFileName = settings.TemplateFirstTimeVisitPatients;

			} else if (reportName.Equals(ReportsInfo.Type.FreeCellsMarketing.ToString())) {
				Type = ReportsInfo.Type.FreeCellsMarketing;
				SqlQuery = settings.MisDbSqlGetFreeCellsMarketing;
				MailTo = settings.MailToFreeCellsMarketing;
				TemplateFileName = settings.TemplateFreeCells;

			} else if (reportName.Equals(ReportsInfo.Type.EmergencyCallsQuantity.ToString())) {
				Type = ReportsInfo.Type.EmergencyCallsQuantity;
				SqlQuery = settings.MisDbSqlGetEmergencyCallsQuantity;
				MailTo = settings.MailToEmergencyCallsQuantity;
				TemplateFileName = settings.TemplateEmergencyCallsQuantity;

			} else if (reportName.Equals(ReportsInfo.Type.RegistryMotivation.ToString())) {
				Type = ReportsInfo.Type.RegistryMotivation;
				SqlQuery = settings.MisDbSqlGetRegistryMotivation;
				MailTo = settings.MailToRegistryMotivation;
				TemplateFileName = settings.TemplateRegistryMotivation;

			} else if (reportName.Equals(ReportsInfo.Type.Reserves.ToString())) {
				Type = ReportsInfo.Type.Reserves;
				SqlQuery = settings.MisDbGetReserves;
				MailTo = settings.MailToReserves;
				TemplateFileName = settings.TemplateReserves;

			} else if (reportName.Equals(ReportsInfo.Type.LicenseEndingDates.ToString())) {
				Type = ReportsInfo.Type.LicenseEndingDates;
				SqlQuery = settings.MisDbGetLicenseEndingDates;
				MailTo = settings.MailToLicenseEndingDates;

			} else if (reportName.Equals(ReportsInfo.Type.Promo.ToString())) {
				Type = ReportsInfo.Type.Promo;
				SqlQuery = settings.VerticaDbSqlGetPromo;
				MailTo = settings.MailToPromo;
				TemplateFileName = settings.TemplatePromo;
				UseVerticaDb = true;

			} else if (reportName.Equals(ReportsInfo.Type.MisTimeSheet.ToString())) {
				Type = ReportsInfo.Type.MisTimeSheet;
				SqlQuery = settings.MisDbSqlGetMisTimeSheet;
				MailTo = settings.MailToMisTimeSheet;
				TemplateFileName = settings.TemplateMisTimeSheet;

			} else if (reportName.Equals(ReportsInfo.Type.PatientsToSha1.ToString())) {
				Type = ReportsInfo.Type.PatientsToSha1;
				SqlQuery = settings.MisDbSqlGetPatientsToSha1;
				TemplateFileName = settings.TemplatePatientsToSha1;

			} else if (reportName.Equals(ReportsInfo.Type.PatientsReferralsDetail.ToString())) {
				Type = ReportsInfo.Type.PatientsReferralsDetail;
				SqlQuery = settings.MisDbSqlGetPatientsReferralsDetail;
				TemplateFileName = settings.TemplatePatientsReferralsDetail;
				MailTo = settings.MailToPatientsReferralsDetail;
				FolderToSave = settings.FolderToSavePatientsReferralsDetail;

			} else if (reportName.Equals(ReportsInfo.Type.FrontOfficeClients.ToString())) {
				Type = ReportsInfo.Type.FrontOfficeClients;
				SqlQuery = settings.MisDbSqlGetFrontOfficeClients;
				TemplateFileName = settings.TemplateFrontOfficeClients;
				MailTo = settings.MailToFrontOfficeClients;

			} else if (reportName.Equals(ReportsInfo.Type.FrontOfficeScheduleRecords.ToString())) {
				Type = ReportsInfo.Type.FrontOfficeScheduleRecords;
				SqlQuery = settings.MisDbSqlGetFrontOfficeScheduleRecords;
				TemplateFileName = settings.TemplateFrontOfficeScheduleRecords;
				MailTo = settings.MailToFrontOfficeScheduleRecords;

			} else if (reportName.Equals(ReportsInfo.Type.FreeCellsToSite.ToString())) {
				Type = ReportsInfo.Type.FreeCellsToSite;
				SqlQuery = settings.MisDbSqlGetFreeCellsToSite;
				TemplateFileName = settings.TemplateFreeCellsToSite;
				MailTo = settings.MailToFreeCellsToSite;
				UploadToServer = true;

			} else if (reportName.Equals(ReportsInfo.Type.FreeCellsToSiteJSON.ToString())) {
				Type = ReportsInfo.Type.FreeCellsToSiteJSON;
				SqlQuery = settings.MisDbSqlGetFreeCellsToSite;
				MailTo = settings.MailToFreeCellsToSite;
				UploadToServer = true;

			} else if (reportName.Equals(ReportsInfo.Type.ScheduleExternalServices.ToString())) {
				Type = ReportsInfo.Type.ScheduleExternalServices;
				SqlQuery = settings.MisDbSqlGetScheduleExternalServices;
				TemplateFileName = settings.TemplateScheduleExternalServices;
				MailTo = settings.MailToScheduleExternalServices;

			} else if (reportName.Equals(ReportsInfo.Type.ServiceListByDoctorsToSite.ToString())) {
				Type = ReportsInfo.Type.ServiceListByDoctorsToSite;
				SqlQuery = settings.VerticaDbSqlGetServiceListByDoctorsToSite;
				TemplateFileName = settings.TemplateServiceListByDoctorsToSite;
				MailTo = settings.MailToServiceListByDoctorsToSite;
				UploadToServer = true;
				UseVerticaDb = true;

			} else if (reportName.Equals(ReportsInfo.Type.ServiceListByDoctorsToSiteJson.ToString())) {
				Type = ReportsInfo.Type.ServiceListByDoctorsToSiteJson;
				SqlQuery = settings.VerticaDbSqlGetServiceListByDoctorsToSite;
				MailTo = settings.MailToServiceListByDoctorsToSite;
				UploadToServer = true;
				UseVerticaDb = true;

			} else if (reportName.Equals(ReportsInfo.Type.RecordCountFrontOffice.ToString())) {
				Type = ReportsInfo.Type.RecordCountFrontOffice;
				SqlQuery = settings.MisDbSqlGetRecordCountFrontOffice;
				TemplateFileName = settings.TemplateRecordCountFrontOffice;
				MailTo = settings.MailToRecordCountFrontOffice;

			} else if (reportName.Equals(ReportsInfo.Type.RFNonResident.ToString())) {
				Type = ReportsInfo.Type.RFNonResident;
				SqlQuery = settings.MisDbSqlGetRFNonResident;
				MailTo = settings.MailToRFNonResident;
				TemplateFileName = settings.TemplateRFNonResident;

			} else if (reportName.Equals(ReportsInfo.Type.Covid19Patients.ToString())) {
				Type = ReportsInfo.Type.Covid19Patients;
				SqlQuery = settings.MisDbSqlGetCovid19Patients;
				MailTo = settings.MailToCovid19Patients;
				TemplateFileName = settings.TemplateCovid19Patients;

			} else if (reportName.Equals(ReportsInfo.Type.ScheduleCallCenter.ToString())) {
				Type = ReportsInfo.Type.ScheduleCallCenter;
				SqlQuery = settings.MisDbSqlGetScheduleCallCenter;
				MailTo = settings.MailToScheduleCallCenter;
				TemplateFileName = settings.TemplateScheduleCallCenter;

			} else if (reportName.Equals(ReportsInfo.Type.AverageCheckRegularMonth.ToString())) {
				Type = ReportsInfo.Type.AverageCheckRegular;
				SqlQuery = settings.MisDbSqlGetAverageCheck;
				TemplateFileName = settings.TemplateAverageCheck;
				MailTo = settings.MailToAverageCheckMonth;

			} else if (reportName.Equals(ReportsInfo.Type.Covid19ByPatientsToGv.ToString())) {
				Type = ReportsInfo.Type.Covid19ByPatientsToGv;
				SqlQuery = settings.MisDbSqlGetCovid19ByPatientsToGv;
				TemplateFileName = settings.TemplateCovid19ByPatientsToGv;
				MailTo = settings.MailToCovid19ByPatientsToGv;

			} else if (reportName.Equals(ReportsInfo.Type.EmployeesCovidTreat.ToString())) {
				Type = ReportsInfo.Type.EmployeesCovidTreat;
				SqlQuery = settings.MisDbSqlGetEmployeesCovidTreat;
				TemplateFileName = settings.TemplateEmployeesCovidTreat;
				MailTo = settings.MailToEmployeesCovidTreat;

			} else if (reportName.Equals(ReportsInfo.Type.PndProviders.ToString())) {
				Type = ReportsInfo.Type.PndProviders;
				SqlQuery = settings.MisDbSqlGetPndProvidersAdult;
				TemplateFileName = settings.TemplatePndProviders;
				MailTo = settings.MailToPndProviders;

			} else if (reportName.Equals(ReportsInfo.Type.ResponsibleForKtKazan.ToString())) {
				Type = ReportsInfo.Type.ResponsibleForKtKazan;
				SqlQuery = settings.MisDbSqlGetResponsibleForKtKazan;
				TemplateFileName = settings.TemplateResponsibleForKtKazan;
				MailTo = settings.MailToResponsibleForKtKazan;

			} else
				IsSettingsLoaded = false;

			if (Type.ToString().Contains("TreatmentsDetails")) {
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails; //settings.VerticaDbSqlGetTreatmentsDetails; //
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;
				UseVerticaDb = false;

				if (Type.ToString().Equals(ReportsInfo.Type.TreatmentsDetailsOther.ToString())) { //Нужно актуализировать запрос к вертике
					UseVerticaDb = false;
					SqlQuery = settings.MisDbSqlGetTreatmentsDetailsOtherIC2; //settings.VerticaDbSqlGetTreatmentsDetailsOtherIC; //
				}
			}

			if (IsSettingsLoaded) {
				Name = ReportsInfo.AcceptedParameters[Type];
				Periodicity = ReportsInfo.Periodicity[Type];
			}
		}

		public void SetMailTo(string mailTo) {
			MailTo = mailTo;
		}

		public void SetPeriod(DateTime dateBegin, DateTime dateEnd) {
			DateBegin = dateBegin;
			DateEnd = dateEnd;
		}
	}
}
