using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports {
	public class ReportsInfo {
		public enum Type {
			FreeCellsDay,
			FreeCellsWeek,
			UnclosedProtocolsWeek,
			UnclosedProtocolsMonth,
			MESUsage,
			OnlineAccountsUsage,
			TelemedicineOnlyIngosstrakh,
			TelemedicineAll,
			NonAppearance,
			VIP_MSSU,
			VIP_Moscow,
			VIP_MSKM,
			VIP_PND,
			RegistryMarks,
			Workload,
			Robocalls,
			UniqueServices,
			UniqueServicesRegions,
			PriceListToSite,
			GBooking,
			PersonalAccountSchedule,
			ProtocolViewCDBSyncEvent,
            FssInfo,
            TimetableBz,
            RecordsFromInsuranceCompanies,
			AverageCheck,
			CompetitiveGroups,
			LicenseStatistics,
			TreatmentsDetails
		};

		public static Dictionary<Type, string> AcceptedParameters = new Dictionary<Type, string> {
			{ Type.FreeCellsDay, "Отчет по свободным слотам" },
			{ Type.FreeCellsWeek, "Отчет по свободным слотам" },
			{ Type.UnclosedProtocolsWeek, "Отчет по неподписанным протоколам" },
			{ Type.UnclosedProtocolsMonth, "Отчет по неподписанным протоколам" },
			{ Type.MESUsage, "Отчет по использованию МЭС" },
			{ Type.OnlineAccountsUsage, "Отчет по записи на прием через личный кабинет" },
			{ Type.TelemedicineOnlyIngosstrakh, "Отчет по приемам телемедицины - только Ингосстрах" },
			{ Type.TelemedicineAll, "Отчет по приемам телемедицины - все типы оплаты" },
			{ Type.NonAppearance, "Отчет по неявкам" },
			{ Type.VIP_MSSU, "Отчет по ВИП-пациентам Сущевка" },
			{ Type.VIP_Moscow, "Отчет по ВИП-пациентам Москва" },
			{ Type.VIP_MSKM, "Отчет по ВИП-пациентам Фрунзенская" },
			{ Type.VIP_PND, "Отчет по ВИП-пациентам ПНД" },
			{ Type.RegistryMarks, "Отчет по оценкам регистратуры" },
			{ Type.Workload, "Отчет по загрузке сотрудников" },
			{ Type.Robocalls, "Информация для автообзвона" },
			{ Type.UniqueServices, "Отчет по уникальным услугам" },
			{ Type.UniqueServicesRegions, "Отчет по уникальным услугам (регионы)" },
			{ Type.PriceListToSite, "Прайс-лист для загрузки на сайт" },
			{ Type.GBooking, "Информация для обзвона пациентов с GBooking" },
			{ Type.PersonalAccountSchedule, "Отчет по записям через личный кабинет" },
			{ Type.ProtocolViewCDBSyncEvent, "Отчет по просмотрам ИБ и синхронизации с ЦБД" },
			{ Type.FssInfo, "Отчет по выданным ЭЛН" },
            { Type.TimetableBz, "Расписание работы врачей" },
            { Type.RecordsFromInsuranceCompanies, "Отчет по записям из страховых компаний" },
            { Type.AverageCheck, "Отчет по среднему чеку" },
			{ Type.CompetitiveGroups, "Отчет по конкурентным группам" },
			{ Type.LicenseStatistics, "Статистика по лицензиям" },
			{ Type.TreatmentsDetails, "Детальный отчет по приемам" }
		};

		public static Dictionary<Type, string> Periodicity = new Dictionary<Type, string> {
			{ Type.FreeCellsDay, "Каждый день в 7:10" },
			{ Type.FreeCellsWeek, "Каждый понедельник в 6:00" },
			{ Type.UnclosedProtocolsWeek, "Каждый понедельник в 10:00" },
			{ Type.UnclosedProtocolsMonth, "Каждый месяц, 2 и 5 числа в 14:50" },
			{ Type.MESUsage, "Каждый понедельник в 6:20" },
			{ Type.OnlineAccountsUsage, "Каждый месяц, 1 числа в 6:00" },
			{ Type.TelemedicineOnlyIngosstrakh, "Каждый понедельник в 7:00" },
			{ Type.TelemedicineAll, "Каждый понедельник в 6:00" },
			{ Type.NonAppearance, "Каждый понедельник в 8:00" },
			{ Type.VIP_MSSU, "Каждый день в 8:02 и 15:02" },
			{ Type.VIP_Moscow, "Каждый день в 8:00 и 15:00" },
			{ Type.VIP_MSKM, "Каждый день в 8:04 и 15:04" },
			{ Type.VIP_PND, "Каждый день в 15:06" },
			{ Type.RegistryMarks, "Каждый понедельник в 7:10" },
			{ Type.Workload, "Каждый месяц, 10 числа в 6:00" },
			{ Type.Robocalls, "Каждый день в 15:00" },
			{ Type.UniqueServices, "Каждый понедельник в 7:57" },
			{ Type.UniqueServicesRegions, "Каждый понедельник в 7:50" },
			{ Type.PriceListToSite, "Каждый день в 2:00" },
			{ Type.GBooking, "Каждый день в 6:10" },
			{ Type.PersonalAccountSchedule, "Каждый день в 10:05" },
			{ Type.ProtocolViewCDBSyncEvent, "Каждый день в 0:00" },
			{ Type.FssInfo, "Каждый понедельник в 11:10" },
			{ Type.TimetableBz, "" },
			{ Type.RecordsFromInsuranceCompanies, "Каждый понедельник в 3:00" },
			{ Type.AverageCheck, "Каждый понедельник в 6:25" },
			{ Type.CompetitiveGroups, "Каждый месяц, 1 числа в 6:40" },
			{ Type.LicenseStatistics, "Каждый день, с 8:00 до 21:00 с интервалом 2 часа" },
			{ Type.TreatmentsDetails, "" }
		};
	}
}
