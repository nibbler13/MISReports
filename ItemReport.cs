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

		public ItemReport(string reportName) {
			if (reportName.Equals(ReportsInfo.Type.FreeCellsDay.ToString())) {
				Type = ReportsInfo.Type.FreeCellsDay;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetFreeCells;
				MailTo = Properties.Settings.Default.MailToFreeCellsDay;
				TemplateFileName = Properties.Settings.Default.TemplateFreeCells;

			} else if (reportName.Equals(ReportsInfo.Type.FreeCellsWeek.ToString())) {
				Type = ReportsInfo.Type.FreeCellsWeek;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetFreeCells;
				MailTo = Properties.Settings.Default.MailToFreeCellsWeek;
				TemplateFileName = Properties.Settings.Default.TemplateFreeCells;

			} else if (reportName.Equals(ReportsInfo.Type.UnclosedProtocolsWeek.ToString())) {
				Type = ReportsInfo.Type.UnclosedProtocolsWeek;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetUnclosedProtocols;
				MailTo = Properties.Settings.Default.MailToUnclosedProtocolsWeek;
				TemplateFileName = Properties.Settings.Default.TemplateUnclosedProtocols;
				FolderToSave = Properties.Settings.Default.FolderToSaveUnclosedProtocols;

			} else if (reportName.Equals(ReportsInfo.Type.UnclosedProtocolsMonth.ToString())) {
				Type = ReportsInfo.Type.UnclosedProtocolsMonth;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetUnclosedProtocols;
				MailTo = Properties.Settings.Default.MailToUnclosedProtocolsMonth;
				TemplateFileName = Properties.Settings.Default.TemplateUnclosedProtocols;
				FolderToSave = Properties.Settings.Default.FolderToSaveUnclosedProtocols;

			} else if (reportName.Equals(ReportsInfo.Type.MESUsage.ToString())) {
				Type = ReportsInfo.Type.MESUsage;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetMESUsage;
				MailTo = Properties.Settings.Default.MailToMESUsage;
				TemplateFileName = Properties.Settings.Default.TemplateMESUsage;
				FolderToSave = Properties.Settings.Default.FolderToSaveMESUsage;

			} else if (reportName.Equals(ReportsInfo.Type.OnlineAccountsUsage.ToString())) {
				Type = ReportsInfo.Type.OnlineAccountsUsage;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetOnlineAccountsUsage;
				MailTo = Properties.Settings.Default.MailToOnlineAccountsUsage;
				TemplateFileName = Properties.Settings.Default.TemplateOnlineAccountsUsage;

			} else if (reportName.Equals(ReportsInfo.Type.TelemedicineOnlyIngosstrakh.ToString())) {
				Type = ReportsInfo.Type.TelemedicineOnlyIngosstrakh;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTelemedicine;
				TemplateFileName = Properties.Settings.Default.TemplateTelemedicine;
				MailTo = Properties.Settings.Default.MailToTelemedicineOnlyIngosstrakh;

			} else if (reportName.Equals(ReportsInfo.Type.TelemedicineAll.ToString())) {
				Type = ReportsInfo.Type.TelemedicineAll;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTelemedicine;
				TemplateFileName = Properties.Settings.Default.TemplateTelemedicine;
				MailTo = Properties.Settings.Default.MailToTelemedicineAll;

			} else if (reportName.Equals(ReportsInfo.Type.NonAppearance.ToString())) {
				Type = ReportsInfo.Type.NonAppearance;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetNonAppearance;
				TemplateFileName = Properties.Settings.Default.TemplateNonAppearance;
				MailTo = Properties.Settings.Default.MailToNonAppearance;
				FolderToSave = Properties.Settings.Default.FolderToSaveNonAppearance;

			} else if (reportName.Equals(ReportsInfo.Type.VIP_MSSU.ToString())) {
				Type = ReportsInfo.Type.VIP_MSSU;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetVIP.Replace("@filialList", "12");
				TemplateFileName = Properties.Settings.Default.TemplateVIP;
				MailTo = Properties.Settings.Default.MailToVIP_MSSU;
				PreviousFile = Properties.Settings.Default.PreviousFileVIP_MSSU;

			} else if (reportName.Equals(ReportsInfo.Type.VIP_Moscow.ToString())) {
				Type = ReportsInfo.Type.VIP_Moscow;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetVIP.Replace("@filialList", "1,5,12,6");
				TemplateFileName = Properties.Settings.Default.TemplateVIP;
				MailTo = Properties.Settings.Default.MailToVIP_Moscow;
				PreviousFile = Properties.Settings.Default.PreviousFileVIP_Moscow;

			} else if (reportName.Equals(ReportsInfo.Type.VIP_MSKM.ToString())) {
				Type = ReportsInfo.Type.VIP_MSKM;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetVIP.Replace("@filialList", "1");
				TemplateFileName = Properties.Settings.Default.TemplateVIP;
				MailTo = Properties.Settings.Default.MailToVIP_MSKM;
				PreviousFile = Properties.Settings.Default.PreviousFileVIP_MSKM;

			} else if (reportName.Equals(ReportsInfo.Type.VIP_PND.ToString())) {
				Type = ReportsInfo.Type.VIP_PND;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetVIP.Replace("@filialList", "6");
				TemplateFileName = Properties.Settings.Default.TemplateVIP;
				MailTo = Properties.Settings.Default.MailToVIP_PND;
				PreviousFile = Properties.Settings.Default.PreviousFileVIP_PND;

			} else if (reportName.Equals(ReportsInfo.Type.RegistryMarks.ToString())) {
				Type = ReportsInfo.Type.RegistryMarks;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetRegistryMarks;
				TemplateFileName = Properties.Settings.Default.TemplateRegistryMarks;
				MailTo = Properties.Settings.Default.MailToRegistryMarks;

			} else if (reportName.Equals(ReportsInfo.Type.Workload.ToString())) {
				Type = ReportsInfo.Type.Workload;
				TemplateFileName = Properties.Settings.Default.TemplateWorkload;
				MailTo = Properties.Settings.Default.MailToWorkload;
				FolderToSave = Properties.Settings.Default.FolderToSaveWorkload;

			} else if (reportName.Equals(ReportsInfo.Type.Robocalls.ToString())) {
				Type = ReportsInfo.Type.Robocalls;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetRobocalls;
				TemplateFileName = Properties.Settings.Default.TemplateRobocalls;
				MailTo = Properties.Settings.Default.MailToRobocalls;

			} else if (reportName.Equals(ReportsInfo.Type.UniqueServices.ToString())) {
				Type = ReportsInfo.Type.UniqueServices;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetUniqueServices;
				TemplateFileName = Properties.Settings.Default.TemplateUniqueServices;
				MailTo = Properties.Settings.Default.MailToUniqueServices;

			} else if (reportName.Equals(ReportsInfo.Type.UniqueServicesRegions.ToString())) {
				Type = ReportsInfo.Type.UniqueServicesRegions;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetUniqueServicesRegions;
				TemplateFileName = Properties.Settings.Default.TemplateUniqueServicesRegions;
				MailTo = Properties.Settings.Default.MailToUniqueServicesRegions;

			} else if (reportName.Equals(ReportsInfo.Type.PriceListToSite.ToString())) {
				Type = ReportsInfo.Type.PriceListToSite;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetPriceListToSite;
				TemplateFileName = Properties.Settings.Default.TemplatePriceListToSite;
				MailTo = Properties.Settings.Default.MailToPriceListToSite;
				FolderToSave = Properties.Settings.Default.FolderToSavePriceListToSite;
				UploadToServer = true;

			} else if (reportName.Equals(ReportsInfo.Type.GBooking.ToString())) {
				Type = ReportsInfo.Type.GBooking;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetGBooking;
				TemplateFileName = Properties.Settings.Default.TemplateGBooking;
				MailTo = Properties.Settings.Default.MailToGBooking;

			} else if (reportName.Equals(ReportsInfo.Type.PersonalAccountSchedule.ToString())) {
				Type = ReportsInfo.Type.PersonalAccountSchedule;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetPersonalAccountSchedule;
				TemplateFileName = Properties.Settings.Default.TemplatePersonalAccountSchedule;
				MailTo = Properties.Settings.Default.MailToPersonalAccountSchedule;

			} else if (reportName.Equals(ReportsInfo.Type.ProtocolViewCDBSyncEvent.ToString())) {
				Type = ReportsInfo.Type.ProtocolViewCDBSyncEvent;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetProtocolViewCDBSyncEvent;
				TemplateFileName = Properties.Settings.Default.TemplateProtocolViewCDBSyncEvent;
				MailTo = Properties.Settings.Default.MailToProtocolViewCDBSyncEvent;
				FolderToSave = Properties.Settings.Default.FolderToSaveProtocolViewCDBSyncEvent;

			} else if (reportName.Equals(ReportsInfo.Type.FssInfo.ToString())) {
				Type = ReportsInfo.Type.FssInfo;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetFssInfo;
				TemplateFileName = Properties.Settings.Default.TemplateFssInfo;
				MailTo = Properties.Settings.Default.MailToFssInfo;

			} else if (reportName.Equals(ReportsInfo.Type.TimetableToProdoctorovRu.ToString())) {
				Type = ReportsInfo.Type.TimetableToProdoctorovRu;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTimetableToProdoctorovRu;
				MailTo = Properties.Settings.Default.MailToTimetableToProdoctorovRu;
				UploadToServer = true;

			} else if (reportName.Equals(ReportsInfo.Type.RecordsFromInsuranceCompanies.ToString())) {
				Type = ReportsInfo.Type.RecordsFromInsuranceCompanies;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetRecordsFromInsuranceCompanies;
				TemplateFileName = Properties.Settings.Default.TemplateRecordsFromInsuranceCompanies;
				MailTo = Properties.Settings.Default.MailToRecordsFromInsuranceCompanies;

			} else if (reportName.Equals(ReportsInfo.Type.AverageCheck.ToString())) {
				Type = ReportsInfo.Type.AverageCheck;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetAverageCheck;
				TemplateFileName = Properties.Settings.Default.TemplateAverageCheck;
				MailTo = Properties.Settings.Default.MailToAverageCheck;

			} else if (reportName.Equals(ReportsInfo.Type.CompetitiveGroups.ToString())) {
				Type = ReportsInfo.Type.CompetitiveGroups;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetCompetitiveGroups;
				TemplateFileName = Properties.Settings.Default.TemplateCompetitiveGroups;
				MailTo = Properties.Settings.Default.MailToAverageCheck;

			} else if (reportName.Equals(ReportsInfo.Type.LicenseStatistics.ToString())) {
				Type = ReportsInfo.Type.LicenseStatistics;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetLicenseStatistics;
				TemplateFileName = Properties.Settings.Default.TemplageLicenseStatistics;
				MailTo = Properties.Settings.Default.MailToLicenseStatistics;
				FolderToSave = Properties.Settings.Default.FolderToSaveLicenseStatistics;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAbsolut.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAbsolut;
				JIDS = "991515382,991519409,991519865,991523030";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAlfa.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAlfa;
				JIDS = "100005,991520911,991514852";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAlliance.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAlliance;
				JIDS = "991511535,991520499,991519440,991521374,991511568";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsBestdoctor.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsBestdoctor;
				JIDS = "991520964, 991526106";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsEnergogarant.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsEnergogarant;
				JIDS = "991523453,991517214,991520348";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhAdult.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhAdult;
				JIDS = "991522348,991522924,991525955,991522442";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsIngosstrakhKid.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsIngosstrakhKid;
				JIDS = "991522386";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

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
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsMetlife.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsMetlife;
				JIDS = "991517927,991523451,991519436";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsOther.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsOther;
				JIDS = "";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsRenessans.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsRenessans;
				JIDS = "991523042,991523280,991523170";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsReso.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsReso;
				JIDS = "991518370,991521272,991523038,991526075,991519595";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsRosgosstrakh.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsRosgosstrakh;
				JIDS = "991511705,1990097479";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSmp.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSmp;
				JIDS = "991516698,991521960";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSogaz.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSogaz;
				JIDS = "991524638";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsSoglasie.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsSoglasie;
				JIDS = "991520913,991518470,991519761,991523028";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsVsk.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsVsk;
				JIDS = "991516556,991520387,991523215,991519361,991525970";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsVtb.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsVtb;
				JIDS = "991515797,991520427";
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetailsAll.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetailsAll;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;
				FolderToSave = Properties.Settings.Default.FolderToSaveTreatmentsDetails;

			} else if (reportName.Equals(ReportsInfo.Type.TimetableToSite.ToString())) {
				Type = ReportsInfo.Type.TimetableToSite;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTimetableToSite;
				MailTo = Properties.Settings.Default.MailToTimetableToSite;
				UploadToServer = true;

			} else if (reportName.Equals(ReportsInfo.Type.MicroSipContactsBook.ToString())) {
				Type = ReportsInfo.Type.MicroSipContactsBook;
				MailTo = Properties.Settings.Default.MailToMicroSipContactsBook;
				FolderToSave = Properties.Settings.Default.FolderToSaveMicroSipContactsBook;

			} else if (reportName.Equals(ReportsInfo.Type.TasksForItilium.ToString())) {
				Type = ReportsInfo.Type.TasksForItilium;
				MailTo = Properties.Settings.Default.MailToTasksForItilium;

			} else if (reportName.Equals(ReportsInfo.Type.FirstTimeVisitPatients.ToString())) {
				Type = ReportsInfo.Type.FirstTimeVisitPatients;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetFirstTimeVisitPatients;
				MailTo = Properties.Settings.Default.MailToFirstTimeVisitPatients;
				TemplateFileName = Properties.Settings.Default.TemplateFirstTimeVisitPatients;

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
