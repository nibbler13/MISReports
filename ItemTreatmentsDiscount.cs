using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports {
    public class ItemTreatmentsDiscount {
        public DateTime DateStart { get; private set; }
        public DateTime? DateEnd { get; private set; }
		public float MainDiscount { get; private set; }
		public List<string> ExcludeDepartments { get; private set; } = new List<string>();
        public List<string> ExcludeKodopers { get; private set; } = new List<string>();
        public Dictionary<Tuple<int, int>, float> DynamicDiscount { get; private set; } = new Dictionary<Tuple<int, int>, float>();
		public bool IsApplyOnlyToServiceList { get; private set; }
		public List<string> ServiceListToApply { get; private set; } = new List<string>();
		public string ApplyToContract { get; private set; }

		public ItemTreatmentsDiscount(DateTime dateStart, DateTime? dateEnd, float mainDiscount, bool isApplyOnlyToServiceList = false, string applyToContract = null) {
			DateStart = dateStart;
			DateEnd = dateEnd;
			MainDiscount = mainDiscount;
			IsApplyOnlyToServiceList = isApplyOnlyToServiceList;
			ApplyToContract = applyToContract;
        }

		public void UpdateMainDiscount(float value) {
			if (MainDiscount != -1)
				return;

			MainDiscount = value;
        }

		public void AddSmpDeptToExclude() {
			ExcludeDepartments.Add("СКОРАЯ МЕДИЦИНСКАЯ ПОМОЩЬ");
		}

		public void AddKtMrtPndSmpDeptToExclude() {
			ExcludeDepartments.AddRange(new List<string> {
				"Компьютерная топография",
				"МРТ",
				"МУЛЬТИСПИРАЛЬНАЯ КОМПЬЮТЕРНАЯ ТОМОГРАФИЯ",
				"ПОМОЩЬ НА ДОМУ",
				"Помощь на дому детское",
				"ЛИЧНЫЙ ВРАЧ",
				"Личный врач детский",
				"СКОРАЯ МЕДИЦИНСКАЯ ПОМОЩЬ"
			});
        }

		public void AddDocOnlineTelemedCovidKodoperToExclude() {
			AddCovid19KodoperToExclude();
			AddDocOnlineKodoperToExclude();
			AddTelemedKodoperToExclude();
        }

        public void AddCovid19KodoperToExclude() {
			ExcludeKodopers.AddRange(
				new List<string> {
					"101898",
					"101899",
					"101935",
					"101936",
					"101937",
					"101938",
					"1002285",
					"1002286",
					"1002287",
					"1002367",
					"1002368",
					"1002369",
					"1002370",
					"1002371",
					"1002372",
					"1002373",
					"1002374",
					"212066",
					"2110635",
					"2110636",
					"2110637",
					"212068",
					"212069",
					"212070",
					"212071",
					"326217",
					"326219",
					"325219",
					"325220",
					"325221",
					"101939",
					"101940",
					"101942",
					"101943",
					"101944",
					"101945"
			});
        }

		public void AddDocOnlineKodoperToExclude() {
			ExcludeKodopers.AddRange(new List<string> {
				"900100",
				"900101",
				"900102",
			});
        }

		public void AddTelemedKodoperToExclude() {
			ExcludeKodopers.Add("900200");
        }

		public void AddCovidInfoToExclude() {
			ExcludeKodopers.Add("900103"); //Информационная поддержка Covid-19)
		}
    }
}
