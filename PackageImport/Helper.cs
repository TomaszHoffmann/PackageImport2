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

        public struct ParametersOfDriveRods
        {
            public string nr_art { get; set; }

            public int handleHeight { get; set; }
            public int dornmass { get; set; }

            public int heightOfDriveRodBox1 { get; set; }
            public int depthOfDriveRodBox1 { get; set; }
            public int thicknessOfDriveRodBox1 { get; set; }
            public int shiftOfDriveRodBox1 { get; set; }

            public int heightOfDriveRodBox2 { get; set; }
            public int depthOfDriveRodBox2 { get; set; }
            public int thicknessOfDriveRodBox2 { get; set; }
            public int shiftOfDriveRodBox2 { get; set; }

            public int heightOfDriveRodBox3 { get; set; }
            public int depthOfDriveRodBox3 { get; set; }
            public int thicknessOfDriveRodBox3 { get; set; }
            public int shiftOfDriveRodBox3 { get; set; }

            public int heightOfDriveRodBox4 { get; set; }
            public int depthOfDriveRodBox4 { get; set; }
            public int thicknessOfDriveRodBox4 { get; set; }
            public int shiftOfDriveRodBox4 { get; set; }

            public int heightOfDriveRodBox5 { get; set; }
            public int depthOfDriveRodBox5 { get; set; }
            public int thicknessOfDriveRodBox5 { get; set; }
            public int shiftOfDriveRodBox5 { get; set; }

            public int heightOfDriveRodBox6 { get; set; }
            public int depthOfDriveRodBox6 { get; set; }
            public int thicknessOfDriveRodBox6 { get; set; }
            public int shiftOfDriveRodBox6 { get; set; }



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
