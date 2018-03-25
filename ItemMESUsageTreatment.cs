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
		public Dictionary<string, int> DictMES { get; set; } = new Dictionary<string, int>(); //0 - Necessary, 1 - ByIndication, 2 - Additional, 3 - ExternalLPU
		public List<string> ListReferralsFromMes { get; set; } = new List<string>();
		public List<string> ListReferralsFromDoc { get; set; } = new List<string>();
		public Dictionary<string, ReferralDetails> DictAllReferrals { get; set; } = new Dictionary<string, ReferralDetails>();
		public string SERVICE_TYPE { get; set; } = string.Empty;
		public string PAYMENT_TYPE { get; set; } = string.Empty;
		public string AGNAME { get; set; } = string.Empty;
		public string AGNUM { get; set; } = string.Empty;

		public class ReferralDetails {
			public string Schid { get; set; } = string.Empty;
			public int IsCompleted { get; set; } = -1;  //0 - incompleted, 1 - completed
			public int RefType { get; set; } = -1; //2 - Lab
		}
	}
}
