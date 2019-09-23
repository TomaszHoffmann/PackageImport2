using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Resources;

namespace PackageImport
{
    public class Helper : IHelper
    {
        public struct ArticlesNhandloInfo
        {
            public string nr_art { get; set; }

            public int przelicz { get; set; }
            public int przlskrz { get; set; }
            public int przelpalet { get; set; }

            public string opisOpak1 { get; set; }
            public string opisOpak2 { get; set; }
            public string opisOpak3 { get; set; }

            public int ok_cel { get; set; }
        }

        ResourceManager resManager = Resources.ProgramResourcesPL.ResourceManager;

        public DataTable GetTable(DataSet dataset, string name)
        {
            Console.WriteLine(resManager.GetString("readingTheWorksheet") + name);

            return dataset.Tables[name];

        }

        public int GetTypeOfTarget(string cel_String)
        {

            int cel;

            switch(cel_String)
            {
                case "ramy":
                    cel = 1;
                    break;
                case "skrzydła":
                    cel = 2;
                    break;
                case "obu":
                    cel = 3;
                    break;
                default:
                    cel = 0;
                    break;
            }

            return cel;

        }

    }
}
