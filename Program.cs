using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MISReports {
	class Program {
		public static string AssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\";

		static void Main(string[] args) {
			SystemLogging.LogMessageToFile("Старт");
			FirebirdClient firebirdClient = new FirebirdClient(
				Properties.Settings.Default.MisDbAddress,
				Properties.Settings.Default.MisDbName,
				Properties.Settings.Default.MisDbUser,
				Properties.Settings.Default.MisDbPassword);

			string dateBegin = "01.01.2018";// DateTime.Now.AddDays(-60).ToShortDateString();
			string dateEnd = "20.02.2018"; //DateTime.Now.AddDays(-30).ToShortDateString();

			//UnclosedProtocols.CreateAndSendReport(firebirdClient, dateBegin, dateEnd);
			MESUsage.CreateAndSendReport(firebirdClient, dateBegin, dateEnd);

			firebirdClient.Close();
			SystemLogging.LogMessageToFile("Завершение работы");
		}
	}
}
