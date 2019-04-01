using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports {
	class ReportsInfo {
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
			ProtocolViewCDBSyncEvent
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
			{ Type.ProtocolViewCDBSyncEvent, "Отчет по просмотрам ИБ и синхронизации с ЦБД" }
		};
	}
}
