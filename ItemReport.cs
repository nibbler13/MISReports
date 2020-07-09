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
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAlfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAlfa;
				JIDS = "100005,991520911,991514852";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAlfaSpb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAlfaSpb;
				JIDS = "990424275"; //80/10-09
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAlliance.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAlliance;
				JIDS = "991511535,991520499,991519440,991521374,991511568";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsBestdoctor.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsBestdoctor;
				JIDS = "991520964, 991526106";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsEnergogarant.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsEnergogarant;
				JIDS = "991523453,991517214,991520348";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhAdult.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhAdult;
				JIDS = "991522348,991522924,991525955,991522442";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhKid.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhKid;
				JIDS = "991522386";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

				//JIDS = "990389345,991511093,991512906,991357338,991370062,991379370,991518734," +
				//	"991518673,991519689,991519376,991520842,991520995,991521326,991522926,991522930," +
				//	"991524101,991524413,991518373,991522932,991524095,991522934,991522938,991516471,991517195," +
				//	"991517199,991518886,991518972,991519618,991520845,991522343,991524107,991524598,991525373," +
				//	"991526193,991514245,991516006,991514230,991517179,991517184,991514243,991526172,991517202," +
				//	"991519343,991519465,991518877,991521414,991519417,991520848,991524071,991522936,991524374," +
				//	"991524084,991514228,991526177,991526091,991517190,991516351,991526130,991518519,991518344," +
				//	"991519305,991520863,991514323,991521813,991523284,991524104,991514234,991526175,991465238," +
				//	"991512358,1990110711,1990108742,1990110725,1990134205,1990138662,1990110728,991510615," +
				//	"991465236,991465233,991465253";

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsLiberty.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsLiberty;
				JIDS = "991517912";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsMetlife.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsMetlife;
				JIDS = "991517927,991523451,991519436";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsOther.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsOther;
				JIDS = "";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsRenessans.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsRenessans;
				JIDS = "991523042,991523280,991523170";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsReso.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsReso;
				JIDS = "991518370,991521272,991523038,991526075,991519595";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsRosgosstrakh.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsRosgosstrakh;
				JIDS = "991511705,1990097479";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSmp.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSmp;
				JIDS = "991516698,991521960";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSogaz.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSogaz;
				JIDS = "991524638";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSoglasie.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSoglasie;
				JIDS = "991520913,991518470,991519761,991523028";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsVsk.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsVsk;
				JIDS = "991516556,991520387,991523215,991519361,991525970";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsVtb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsVtb;
				JIDS = "991515797,991520427";
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAll.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAll;
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhSochi.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhSochi;
				JIDS = "991512906"; //4986881-19/16
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhKrasnodar.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhKrasnodar;
				JIDS = "991357338"; //№ 567751-19/11
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhUfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhUfa;
				JIDS = "991370062"; //№ 681187-19/11
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhSpb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhSpb;
				JIDS = "990389345"; //267673-19/09
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhKazan.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhKazan;
				JIDS = "991379370"; //№ 714760-19/11
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsBestDoctorSpb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsBestDoctorSpb;
				JIDS = "991523486"; //522-78-2018
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsBestDoctorUfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsBestDoctorUfa;
				JIDS = "991523489"; //535-02-18
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSogazUfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSogazUfa;
				JIDS = "991524671,991524697"; //2719RP055, ДС № 2719RP055-02  к дог. №2719RP055 (ГК «БАШНЕФТЬ»)
				SqlQuery = settings.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = settings.TemplateTreatmentsDetails;
				MailTo = settings.MailToTreatmentsDetails;
				FolderToSave = settings.FolderToSaveTreatmentsDetails;
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
				SqlQuery = settings.VerticaDbSqlQueryGetPromo;
				MailTo = settings.MailToPromo;
				TemplateFileName = settings.TemplatePromo;
				UseVerticaDb = true;

			} else
				IsSettingsLoaded = false;

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
