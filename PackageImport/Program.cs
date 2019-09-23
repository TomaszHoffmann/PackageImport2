using Excel;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Resources;

namespace PackageImport
{
    public class Program
    {
        private string _excelPath, _databaseConnectionString;

        public Program (string excelPath, string databaseConnectionString)
        {
            _excelPath = excelPath;
            _databaseConnectionString = databaseConnectionString;
        }

        const int NR_ART_LEN = 25;

        IHelper helper = new Helper();

        List<Helper.ArticlesNhandloInfo> articlesNhandlo = new List<Helper.ArticlesNhandloInfo>();

        ResourceManager resManager = Resources.ProgramResourcesPL.ResourceManager ;
       
        private bool ParseExcelFile(Stream stream)
        {
            
            bool failed = false;

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();
            excelReader.Close();
            
           

            DataTable table = helper.GetTable(result, "Sheet1");

            int count = 0;
            
            foreach (DataRow row in table.Rows)
            {
                
                count += 1;
                if (count > 1)
                {

                    Helper.ArticlesNhandloInfo ArticlesInfo = new Helper.ArticlesNhandloInfo()
                    {

                    nr_art = Convert.ToString(row["Materiał"]),

                    przelicz = Convert.ToInt32(row["Ilość opakowanie 1"]),
                    przlskrz = Convert.ToInt32(row["Ilość opakowanie 2"]),
                    przelpalet = Convert.ToInt32(row["Ilość opakowanie 3"]),


                    opisOpak1 = Convert.ToString(row["Opis opakowanie 1"]),
                    opisOpak2 = Convert.ToString(row["Opis Opakowanie 2"]),
                    opisOpak3 = Convert.ToString(row["Opis Opakowanie 3"]),

                    ok_cel = helper.GetTypeOfTarget(Convert.ToString(row["Mocowany Do"]))

                    };

                    articlesNhandlo.Add(ArticlesInfo);
                    

                }

            }


            return !failed;
        }

        private void LoadExcel(string excelPath)
        {
            bool ok = false;


            FileStream stream;
            try
            {
                stream = File.Open(excelPath, FileMode.Open, FileAccess.Read);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(resManager.GetString("ImportAborted"));
                Console.ReadKey();
                return;
            }

            try
            {
                ok = ParseExcelFile(stream);
                Console.WriteLine(resManager.GetString("ok"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

          

            

        }

        private void SaveNhandloParameters(string databaseConnectionString)
        {
           
            using (SqlConnection con = new SqlConnection(databaseConnectionString))
            using (SqlCommand command = con.CreateCommand())
            {
                con.Open();
                command.Parameters.Add("@NR_ART", SqlDbType.NVarChar);
                command.Parameters.Add("@opak1", SqlDbType.Int);
                command.Parameters.Add("@opak2", SqlDbType.Int);
                command.Parameters.Add("@opak3", SqlDbType.Int);
                command.Parameters.Add("@opisOpak1", SqlDbType.NVarChar);
                command.Parameters.Add("@opisOpak2", SqlDbType.NVarChar);
                command.Parameters.Add("@opisOpak3", SqlDbType.NVarChar);
                command.Parameters.Add("@ok_cel", SqlDbType.Int);


                foreach (var item in articlesNhandlo)
                {

                    command.CommandText = "UPDATE DBO.NHANDLOTABLE SET przelicz=@opak1, PRZELSKRZ=@opak2, PRZELPALET=@opak3, ";
                    command.CommandText += "opisOpak1 = @opisOpak1, opisOpak2 = @opisOpak2, opisOpak3 = @opisOpak3 , ok_cel=@ok_cel ";
                    command.CommandText += " WHERE NR_ART=@NR_ART";
                    
                    command.Parameters["@NR_ART"].Value = item.nr_art;

                    command.Parameters["@opak1"].Value = item.przelicz; 
                    command.Parameters["@opak2"].Value = item.przlskrz; 
                    command.Parameters["@opak3"].Value = item.przelpalet; 
                    command.Parameters["@opisOpak1"].Value = item.opisOpak1; 
                    command.Parameters["@opisOpak2"].Value = item.opisOpak2; 
                    command.Parameters["@opisOpak3"].Value = item.opisOpak3;
                    command.Parameters["@ok_cel"].Value = item.ok_cel;

                    Console.WriteLine(resManager.GetString("Loading" ) + " nhandlo");

                    try
                    {
                        command.ExecuteNonQuery();
                      
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        Console.ReadKey();

                    }

                }
                con.Close();
                Console.WriteLine("Press any key to end");
                Console.ReadKey();
            }

            }

        public static void Main(string[] args)

        {
            Program program = new Program(@"C:\Users\tomasz.hoffmann\Desktop\LPT 17072019.xlsx", @"Data Source=.\sqlexpress;Initial Catalog=BMW5.6.2.1; Integrated Security=False;User ID=sa;Password=Whokna123@");
            program.LoadExcel(program._excelPath);
            program.SaveNhandloParameters(program._databaseConnectionString);
        }
        

       /* static void Main(string[] args)
        {
            LoadExcelMain(args[0], args[1]);
        }*/
    }

}
