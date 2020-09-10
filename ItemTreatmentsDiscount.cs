﻿using System;
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

		public ItemTreatmentsDiscount(DateTime dateStart, DateTime? dateEnd, float mainDiscount) {
			DateStart = dateStart;
			DateEnd = dateEnd;
			MainDiscount = mainDiscount;
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
					"326217",
					"325219",
					"101939",
					"101940"
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
    }
}