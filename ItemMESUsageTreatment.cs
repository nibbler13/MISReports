using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports {
	public class ItemMESUsageTreatment {
		public string TREATDATE { get; set; } = string.Empty;
		public string FILIAL { get; set; } = string.Empty;
		public string DEPNAME { get; set; } = string.Empty;
		public string DOCNAME { get; set; } = string.Empty;
		public string HISTNUM { get; set; } = string.Empty;
		public string CLIENTNAME { get; set; } = string.Empty;
		public string MKBCODE { get; set; } = string.Empty;
		public string AGE { get; set; } = string.Empty;
		public List<string> ListMES { get; set; } = new List<string>();
		public List<string> ListReferralsFromMes { get; set; } = new List<string>();
		public List<string> ListReferralsFromDoc { get; set; } = new List<string>();
		public Dictionary<string, int> ListAllReferrals { get; set; } = new Dictionary<string, int>();
	}
}
