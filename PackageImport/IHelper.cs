using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PackageImport
{
    interface IHelper
    {

        DataTable GetTable(DataSet dataset, string name);

        int GetTypeOfTarget(string cel_String);
    }
}
