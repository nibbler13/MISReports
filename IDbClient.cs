using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISReports {
    interface IDbClient {
        DataTable GetDataTable(string query, Dictionary<string, object> parameters = null);
        bool ExecuteUpdateQuery(string query, Dictionary<string, object> parameters);
        void Close();
    }
}
