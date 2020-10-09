﻿using System;
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
			MESUsageFull,
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
            TimetableToProdoctorovRu,
            RecordsFromInsuranceCompanies,
			AverageCheckRegular,
			AverageCheckIGS,
			AverageCheckMSK,
			AverageCheckCash,
			CompetitiveGroups,
			LicenseStatistics,
			TreatmentsDetailsAbsolut,
			TreatmentsDetailsAlfa,
			TreatmentsDetailsAlfaSpb,
			TreatmentsDetailsAlliance,
			TreatmentsDetailsBestdoctor,
			TreatmentsDetailsVsk,
			TreatmentsDetailsVtb,
			TreatmentsDetailsOther,
			TreatmentsDetailsIngosstrakhAdult,
			TreatmentsDetailsIngosstrakhKid,
			TreatmentsDetailsLiberty,
			TreatmentsDetailsMetlife,
			TreatmentsDetailsRosgosstrakh,
			TreatmentsDetailsRenessans,
			TreatmentsDetailsReso,
			TreatmentsDetailsSmp,
			TreatmentsDetailsSogaz,
			TreatmentsDetailsSoglasie,
			TreatmentsDetailsEnergogarant,
			TreatmentsDetailsAll,
			TreatmentsDetailsIngosstrakhSochi,
			TreatmentsDetailsIngosstrakhKrasnodar,
			TreatmentsDetailsIngosstrakhUfa,
			TreatmentsDetailsIngosstrakhSpb,
			TreatmentsDetailsIngosstrakhKazan,
			TreatmentsDetailsBestDoctorSpb,
			TreatmentsDetailsBestDoctorUfa,
			TreatmentsDetailsSogazUfa,
			TimetableToSite,
			MicroSipContactsBook,
			TasksForItilium,
			FirstTimeVisitPatients,
			FreeCellsMarketing,
			EmergencyCallsQuantity,
			RegistryMotivation,
			Reserves,
			LicenseEndingDates,
			Promo,
			MisTimeSheet,
			PatientsToSha1,
			PatientsReferralsDetail,
			FrontOfficeClients
		};

		public static Dictionary<Type, string> AcceptedParameters = new Dictionary<Type, string> {
			{ Type.FreeCellsDay, "Отчет по свободным слотам" },
			{ Type.FreeCellsWeek, "Отчет по свободным слотам" },
			{ Type.UnclosedProtocolsWeek, "Отчет по неподписанным протоколам" },
			{ Type.UnclosedProtocolsMonth, "Отчет по неподписанным протоколам" },
			{ Type.MESUsage, "Отчет по использованию МЭС" },
			{ Type.MESUsageFull, "Отчет по использованию МЭС (полный)" },
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
			{ Type.PriceListToSite, "Прайс-лист для загрузки на сайт klinikabudzdorov.ru" },
			{ Type.GBooking, "Информация для обзвона пациентов с GBooking" },
			{ Type.PersonalAccountSchedule, "Отчет по записям через личный кабинет" },
			{ Type.ProtocolViewCDBSyncEvent, "Отчет по просмотрам ИБ и синхронизации с ЦБД" },
			{ Type.FssInfo, "Отчет по выданным ЭЛН" },
            { Type.TimetableToProdoctorovRu, "Выгрузка расписания на сайт prodoctorov.ru" },
            { Type.RecordsFromInsuranceCompanies, "Отчет по записям из страховых компаний" },
            { Type.AverageCheckRegular, "Отчет по среднему чеку" },
            { Type.AverageCheckIGS, "Отчет по среднему чеку (ИГС)" },
            { Type.AverageCheckMSK, "Отчет по среднему чеку (МСК)" },
            { Type.AverageCheckCash, "Отчет по среднему чеку (Физики-факт)" },
			{ Type.CompetitiveGroups, "Отчет по конкурентным группам" },
			{ Type.LicenseStatistics, "Статистика по лицензиям" },
			{ Type.TreatmentsDetailsAbsolut, "Детальный отчет по приемам АбсолютСтрахование" },
			{ Type.TreatmentsDetailsAlfa, "Детальный отчет по приемам Альфастрахование" },
			{ Type.TreatmentsDetailsAlfaSpb, "Детальный отчет по приемам Альфастрахование Санкт-Петербург" },
			{ Type.TreatmentsDetailsAlliance, "Детальный отчет по приемам Альянс" },
			{ Type.TreatmentsDetailsBestdoctor, "Детальный отчет по приемам БестДоктор" },
			{ Type.TreatmentsDetailsVsk, "Детальный отчет по приемам ВСК" },
			{ Type.TreatmentsDetailsVtb, "Детальный отчет по приемам ВТБ" },
			{ Type.TreatmentsDetailsOther, "Детальный отчет по приемам Другие СК" },
			{ Type.TreatmentsDetailsIngosstrakhAdult, "Детальный отчет по приемам Ингосстрах взр" },
			{ Type.TreatmentsDetailsIngosstrakhKid, "Детальный отчет по приемам Ингосстрах дет" },
			{ Type.TreatmentsDetailsLiberty, "Детальный отчет по приемам ЛибертиСтрахование" },
			{ Type.TreatmentsDetailsMetlife, "Детальный отчет по приемам Метлайф" },
			{ Type.TreatmentsDetailsRosgosstrakh, "Детальный отчет по приемам Росгосстрах" },
			{ Type.TreatmentsDetailsRenessans, "Детальный отчет по приемам СК Ренессанс" },
			{ Type.TreatmentsDetailsReso, "Детальный отчет по приемам СК РЕСО" },
			{ Type.TreatmentsDetailsSmp, "Детальный отчет по приемам СМП страхование" },
			{ Type.TreatmentsDetailsSogaz, "Детальный отчет по приемам СОГАЗ" },
			{ Type.TreatmentsDetailsSoglasie, "Детальный отчет по приемам Согласие" },
			{ Type.TreatmentsDetailsEnergogarant, "Детальный отчет по приемам Энергогарант" },
			{ Type.TreatmentsDetailsAll, "Детальный отчет по приемам" },
			{ Type.TreatmentsDetailsIngosstrakhSochi, "Детальный отчет по приемам Ингосстрах Сочи" },
			{ Type.TreatmentsDetailsIngosstrakhKrasnodar, "Детальный отчет по приемам Ингосстрах Краснодар" },
			{ Type.TreatmentsDetailsIngosstrakhUfa, "Детальный отчет по приемам Ингосстрах Уфа" },
			{ Type.TreatmentsDetailsIngosstrakhSpb, "Детальный отчет по приемам Ингосстрах Санкт-Петербург" },
			{ Type.TreatmentsDetailsIngosstrakhKazan, "Детальный отчет по приемам Ингосстрах Казань" },
			{ Type.TreatmentsDetailsBestDoctorSpb, "Детальный отчет по приемам БестДоктор Санкт-Петербург" },
			{ Type.TreatmentsDetailsBestDoctorUfa, "Детальный отчет по приемам БестДоктор Уфа" },
			{ Type.TreatmentsDetailsSogazUfa, "Детальный отчет по приемам СОГАЗ Уфа" },
			{ Type.TimetableToSite, "Выгрузка расписания на сайт klinikabudzdorov.ru" },
			{ Type.MicroSipContactsBook, "Справочик контактов для MicroSip" },
			{ Type.TasksForItilium, "Задачи на январь 2020" },
			{ Type.FirstTimeVisitPatients, "Отчет по первичным пациентам" },
			{ Type.FreeCellsMarketing, "Отчет по свободным слотам" },
			{ Type.EmergencyCallsQuantity, "Отчет по количеству вызовов СМП" },
			{ Type.RegistryMotivation, "Расчет мотивации для регистратуры" },
			{ Type.Reserves, "Отчет по резервам в расписании" },
			{ Type.LicenseEndingDates, "Статистика по окончанию действия лицензий" },
			{ Type.Promo, "Список действующих акций" },
			{ Type.MisTimeSheet, "Табель из МИС" },
			{ Type.PatientsToSha1, "Получение хэш-суммы SHA1 для списка пациентов" },
			{ Type.PatientsReferralsDetail, "Отчет по направлениям пациентов (фронт-офис)" },
			{ Type.FrontOfficeClients, "Список наличных пациентов на сегодня (фронт-офис)" }
		};

		public static Dictionary<Type, string> Periodicity = new Dictionary<Type, string> {
			{ Type.FreeCellsDay, "Каждый день в 7:10" },
			{ Type.FreeCellsWeek, "Каждый понедельник в 6:00" },
			{ Type.UnclosedProtocolsWeek, "Каждый понедельник в 10:00" },
			{ Type.UnclosedProtocolsMonth, "Каждый месяц, 2 и 5 числа в 14:50" },
			{ Type.MESUsage, "Каждый понедельник в 6:20" },
			{ Type.MESUsageFull, "" },
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
			{ Type.TimetableToProdoctorovRu, "" },
			{ Type.RecordsFromInsuranceCompanies, "Каждый понедельник в 3:00" },
			{ Type.AverageCheckRegular, "Каждый понедельник в 6:25" },
			{ Type.AverageCheckIGS, "Каждый день в 7:30" },
			{ Type.AverageCheckMSK, "Каждый понедельник и каждое 1 число месяца в 6:40" },
			{ Type.AverageCheckCash, "" },
			{ Type.CompetitiveGroups, "Каждый месяц, 1 числа в 6:40" },
			{ Type.LicenseStatistics, "Каждый день, с 8:00 до 21:00 с интервалом 2 часа" },
			{ Type.TreatmentsDetailsAbsolut, "" },
			{ Type.TreatmentsDetailsAlfa, "" },
			{ Type.TreatmentsDetailsAlfaSpb, "" },
			{ Type.TreatmentsDetailsAlliance, "" },
			{ Type.TreatmentsDetailsBestdoctor, "" },
			{ Type.TreatmentsDetailsVsk, "" },
			{ Type.TreatmentsDetailsVtb, "" },
			{ Type.TreatmentsDetailsOther, "" },
			{ Type.TreatmentsDetailsIngosstrakhAdult, "" },
			{ Type.TreatmentsDetailsIngosstrakhKid, "" },
			{ Type.TreatmentsDetailsLiberty, "" },
			{ Type.TreatmentsDetailsMetlife, "" },
			{ Type.TreatmentsDetailsRosgosstrakh, "" },
			{ Type.TreatmentsDetailsRenessans, "" },
			{ Type.TreatmentsDetailsReso, "" },
			{ Type.TreatmentsDetailsSmp, "" },
			{ Type.TreatmentsDetailsSogaz, "" },
			{ Type.TreatmentsDetailsSoglasie, "" },
			{ Type.TreatmentsDetailsEnergogarant, "" },
			{ Type.TreatmentsDetailsAll, "Каждый четверг в 7:00" },
			{ Type.TreatmentsDetailsIngosstrakhSochi, "" },
			{ Type.TreatmentsDetailsIngosstrakhKrasnodar, "" },
			{ Type.TreatmentsDetailsIngosstrakhUfa, "" },
			{ Type.TreatmentsDetailsIngosstrakhSpb, "" },
			{ Type.TreatmentsDetailsIngosstrakhKazan, "" },
			{ Type.TreatmentsDetailsBestDoctorSpb, "" },
			{ Type.TreatmentsDetailsBestDoctorUfa, "" },
			{ Type.TreatmentsDetailsSogazUfa, "" },
			{ Type.TimetableToSite, "" },
			{ Type.MicroSipContactsBook, "Каждый день в 1:00" },
			{ Type.TasksForItilium, "Каждый день в 9:00" },
			{ Type.FirstTimeVisitPatients, "Каждый понедельник в 5:00, каждый месяц" },
			{ Type.FreeCellsMarketing, "Каждый день в 7:20 и каждый понедельник в 6:10" },
			{ Type.EmergencyCallsQuantity, "Каждый понедельник в 2:20 и каждое 1 число в 2:30" },
			{ Type.RegistryMotivation, "Каждый месяц 10 числа в 3:00" },
			{ Type.Reserves, "Каждый месяц 3 и 18 числа в в 3:10" },
			{ Type.LicenseEndingDates, "Каждый день в 5:00" },
			{ Type.Promo, "Каждый понедельник" },
			{ Type.MisTimeSheet, "Каждое 1 и 16 число месяца" },
			{ Type.PatientsToSha1, "" },
			{ Type.PatientsReferralsDetail, "" },
			{ Type.FrontOfficeClients, "Каждый день в 7:30" }
		};
	}
}
