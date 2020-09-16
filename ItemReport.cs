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

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAlfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAlfa;
				JIDS = "100005,991520911,991514852";

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2018, 6, 1), null, 20);
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 6, 1), new DateTime(2020, 6, 30), 2.9f);
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAlfaSpb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAlfaSpb;
				JIDS = "990424275"; //80/10-09

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAlliance.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAlliance;
				JIDS = "991511535,991520499,991519440,991521374,991511568";

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
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

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsEnergogarant.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsEnergogarant;
				JIDS = "991523453,991517214,991520348";

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhAdult.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhAdult;
				JIDS = "991522348,991522924,991525955,991522442";

				//=====================
				//Управленческая скидка
				//=====================
				//ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 20);
				//discount.ExcludeDepartments.Add("СКОРАЯ МЕДИЦИНСКАЯ ПОМОЩЬ");
				//TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 10);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhKid.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhKid;
				JIDS = "991522386";

				//=====================
				//Управленческая скидка
				//=====================
				//ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 20);
				//discount.ExcludeDepartments.Add("СКОРАЯ МЕДИЦИНСКАЯ ПОМОЩЬ");
				//TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 10);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsLiberty.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsLiberty;
				JIDS = "991517912";

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsMetlife.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsMetlife;
				JIDS = "991517927,991523451,991519436";

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

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsOther.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsOther;
				JIDS = "";

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsRenessans.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsRenessans;
				JIDS = "991523042,991523280,991523170";

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 10);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsReso.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsReso;
				JIDS = "991518370,991521272,991523038,991526075,991519595";

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), -1);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 10);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsRosgosstrakh.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsRosgosstrakh;
				JIDS = "991511705,1990097479";

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 10);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSmp.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSmp;
				JIDS = "991516698,991521960";

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSogaz.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSogaz;
				JIDS = "991524638";

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(500000, 1000000), 5);
				discount.DynamicDiscount.Add(new Tuple<int, int>(1000001, 2000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(2000001, 3000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(3000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 2, 1), new DateTime(2020, 12, 31), 10);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSoglasie.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSoglasie;
				JIDS = "991520913,991518470,991519761,991523028";

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

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsVsk.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsVsk;
				JIDS = "991516556,991520387,991523215,991519361,991525970";

				ItemTreatmentsDiscount discount = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), -1);
				discount.DynamicDiscount.Add(new Tuple<int, int>(4000000, 5000000), 10);
				discount.DynamicDiscount.Add(new Tuple<int, int>(5000001, 10000000), 15);
				discount.DynamicDiscount.Add(new Tuple<int, int>(10000001, -1), 20);
				discount.AddSmpDeptToExclude();
				discount.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount);

				ItemTreatmentsDiscount discount2 = new ItemTreatmentsDiscount(new DateTime(2020, 1, 1), new DateTime(2020, 12, 31), 5);
				discount2.AddKtMrtPndSmpDeptToExclude();
				discount2.AddDocOnlineTelemedCovidKodoperToExclude();
				TreatmentsDiscounts.Add(discount2);

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsVtb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsVtb;
				JIDS = "991515797,991520427";

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAll.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAll;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhSochi.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhSochi;
				JIDS = "991512906"; //4986881-19/16

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhKrasnodar.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhKrasnodar;
				JIDS = "991357338"; //№ 567751-19/11

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhUfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhUfa;
				JIDS = "991370062"; //№ 681187-19/11

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhSpb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhSpb;
				JIDS = "990389345"; //267673-19/09

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhKazan.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhKazan;
				JIDS = "991379370"; //№ 714760-19/11

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsBestDoctorSpb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsBestDoctorSpb;
				JIDS = "991523486"; //522-78-2018

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsBestDoctorUfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsBestDoctorUfa;
				JIDS = "991523489"; //535-02-18

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSogazUfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSogazUfa;
				JIDS = "991524671,991524697"; //2719RP055, ДС № 2719RP055-02  к дог. №2719RP055 (ГК «БАШНЕФТЬ»)
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

			} else
				IsSettingsLoaded = false;

			if (Type.ToString().Contains("TreatmentsDetails")) {
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails; //settings.VerticaDbSqlGetTreatmentsDetails; //
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;
				UseVerticaDb = false;

				if (Type.ToString().Equals(ReportsInfo.Type.TreatmentsDetailsOther.ToString()))
					SqlQuery = settings.MisDbSqlGetTreatmentsDetailsOtherIC;
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
