using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports {
	public class ItemReport {
		public ReportsInfo.Type Type { get; private set; }
		public string Name { get; private set; }
		public string SqlQuery { get; private set; }
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

			} else if (reportName.Equals(ReportsInfo.Type.TimetableBz.ToString())) {
				Type = ReportsInfo.Type.TimetableBz;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTimetableBz;
				TemplateFileName = Properties.Settings.Default.TemplateTimetableBz;
				MailTo = Properties.Settings.Default.MailToTimetableBz;
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
				SqlQuery = Properties.Settings.Default.MisDbSqlGetAverageCheck;
				TemplateFileName = Properties.Settings.Default.TemplateCompetitiveGroups;
				MailTo = Properties.Settings.Default.MailToAverageCheck;

			} else if (reportName.Equals(ReportsInfo.Type.LicenseStatistics.ToString())) {
				Type = ReportsInfo.Type.LicenseStatistics;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetLicenseStatistics;
				TemplateFileName = Properties.Settings.Default.TemplageLicenseStatistics;
				MailTo = Properties.Settings.Default.MailToLicenseStatistics;
				FolderToSave = Properties.Settings.Default.FolderToSaveLicenseStatistics;

			} else if (reportName.Equals(ReportsInfo.Type.TreatmentsDetails.ToString())) {
				Type = ReportsInfo.Type.TreatmentsDetails;
				SqlQuery = Properties.Settings.Default.MisDbSqlGetTreatmentsDetails;
				TemplateFileName = Properties.Settings.Default.TemplateTreatmentsDetails;
				MailTo = Properties.Settings.Default.MailToTreatmentsDetails;

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
