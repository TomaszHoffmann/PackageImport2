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
        List<Helper.ParametersOfDriveRods> parametersOfDriveRods = new List<Helper.ParametersOfDriveRods>();

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

                    Helper.ParametersOfDriveRods parametersOfDriveRodsInfo = new Helper.ParametersOfDriveRods()
                    {
                        nr_art = Convert.ToString(row["Materiał"]),

                        handleHeight = Convert.ToInt32(row["Wysokość Klamki"]),
                        dornmass = Convert.ToInt32(row["Dornmass"]),

                        heightOfDriveRodBox1 = Convert.ToInt32(row["Puszka Zasuw. 1 Wysokość"]),
                        depthOfDriveRodBox1 = Convert.ToInt32(row["Puszka Zasuw. 1 Głębokość"]),
                        thicknessOfDriveRodBox1 = Convert.ToInt32(row["Puszka Zasuw. 1 Grubość"]),
                        shiftOfDriveRodBox1 = Convert.ToInt32(row["Puszka Zasuw. 1 Grubość"]),

                        heightOfDriveRodBox2 = Convert.ToInt32(row["Puszka Zasuw. 2 Wysokość"]),
                        depthOfDriveRodBox2 = Convert.ToInt32(row["Puszka Zasuw. 2 Głębokość"]),
                        thicknessOfDriveRodBox2 = Convert.ToInt32(row["Puszka Zasuw. 2 Grubość"]),
                        shiftOfDriveRodBox2 = Convert.ToInt32(row["Puszka Zasuw. 2 Grubość"]),

                        heightOfDriveRodBox3 = Convert.ToInt32(row["Puszka Zasuw. 3 Wysokość"]),
                        depthOfDriveRodBox3 = Convert.ToInt32(row["Puszka Zasuw. 3 Głębokość"]),
                        thicknessOfDriveRodBox3 = Convert.ToInt32(row["Puszka Zasuw. 3 Grubość"]),
                        shiftOfDriveRodBox3 = Convert.ToInt32(row["Puszka Zasuw. 3 Grubość"]),

                        heightOfDriveRodBox4 = Convert.ToInt32(row["Puszka Zasuw. 4 Wysokość"]),
                        depthOfDriveRodBox4 = Convert.ToInt32(row["Puszka Zasuw. 4 Głębokość"]),
                        thicknessOfDriveRodBox4 = Convert.ToInt32(row["Puszka Zasuw. 4 Grubość"]),
                        shiftOfDriveRodBox4 = Convert.ToInt32(row["Puszka Zasuw. 4 Grubość"]),

                        heightOfDriveRodBox5 = Convert.ToInt32(row["Puszka Zasuw. 5 Wysokość"]),
                        depthOfDriveRodBox5 = Convert.ToInt32(row["Puszka Zasuw. 5 Głębokość"]),
                        thicknessOfDriveRodBox5 = Convert.ToInt32(row["Puszka Zasuw. 5 Grubość"]),
                        shiftOfDriveRodBox5 = Convert.ToInt32(row["Puszka Zasuw. 5 Grubość"]),

                        heightOfDriveRodBox6 = Convert.ToInt32(row["Puszka Zasuw. 6 Wysokość"]),
                        depthOfDriveRodBox6 = Convert.ToInt32(row["Puszka Zasuw. 6 Głębokość"]),
                        thicknessOfDriveRodBox6 = Convert.ToInt32(row["Puszka Zasuw. 6 Grubość"]),
                        shiftOfDriveRodBox6 = Convert.ToInt32(row["Puszka Zasuw. 6 Grubość"])

                    };

                    parametersOfDriveRods.Add(parametersOfDriveRodsInfo);
                    

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

                Console.WriteLine(resManager.GetString("Loading" ) + " nhandlo");
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

        private void SaveDriveRodsParameters(string databaseConnectionString)
        {

            using (SqlConnection con = new SqlConnection(databaseConnectionString))
            using (SqlCommand command = con.CreateCommand())
            {
                con.Open();
                command.Parameters.Add("@RowCount", SqlDbType.Int).Direction = ParameterDirection.Output;
                command.Parameters.Add("@NR_ART", SqlDbType.NVarChar);
                command.Parameters.Add("@handleHeight", SqlDbType.Int);
                command.Parameters.Add("@dornmass", SqlDbType.Int);
                
                command.Parameters.Add("@heightOfDriveRodBox1", SqlDbType.Int);
                command.Parameters.Add("@depthOfDriveRodBox1", SqlDbType.Int);
                command.Parameters.Add("@thicknessOfDriveRodBox1", SqlDbType.Int);
                command.Parameters.Add("@shiftOfDriveRodBox1", SqlDbType.Int);

                command.Parameters.Add("@heightOfDriveRodBox2", SqlDbType.Int);
                command.Parameters.Add("@depthOfDriveRodBox2", SqlDbType.Int);
                command.Parameters.Add("@thicknessOfDriveRodBox2", SqlDbType.Int);
                command.Parameters.Add("@shiftOfDriveRodBox2", SqlDbType.Int);

                command.Parameters.Add("@heightOfDriveRodBox3", SqlDbType.Int);
                command.Parameters.Add("@depthOfDriveRodBox3", SqlDbType.Int);
                command.Parameters.Add("@thicknessOfDriveRodBox3", SqlDbType.Int);
                command.Parameters.Add("@shiftOfDriveRodBox3", SqlDbType.Int);

                command.Parameters.Add("@heightOfDriveRodBox4", SqlDbType.Int);
                command.Parameters.Add("@depthOfDriveRodBox4", SqlDbType.Int);
                command.Parameters.Add("@thicknessOfDriveRodBox4", SqlDbType.Int);
                command.Parameters.Add("@shiftOfDriveRodBox4", SqlDbType.Int);

                command.Parameters.Add("@heightOfDriveRodBox5", SqlDbType.Int);
                command.Parameters.Add("@depthOfDriveRodBox5", SqlDbType.Int);
                command.Parameters.Add("@thicknessOfDriveRodBox5", SqlDbType.Int);
                command.Parameters.Add("@shiftOfDriveRodBox5", SqlDbType.Int);

                command.Parameters.Add("@heightOfDriveRodBox6", SqlDbType.Int);
                command.Parameters.Add("@depthOfDriveRodBox6", SqlDbType.Int);
                command.Parameters.Add("@thicknessOfDriveRodBox6", SqlDbType.Int);
                command.Parameters.Add("@shiftOfDriveRodBox6", SqlDbType.Int);


                Console.WriteLine(resManager.GetString("Loading") + " zasow");
                foreach (var item in parametersOfDriveRods)
                {

                    command.CommandText = "UPDATE DBO.ZASOW SET wys_klam = @handleHeight, szerpuszki=@thicknessOfDriveRodBox1, glebpuszki = @depthOfDriveRodBox1,  ";
                    command.CommandText += "polozenie2=@heightOfDriveRodBox2, szerpuszk2=@thicknessOfDriveRodBox2, glebpuszk2 = @depthOfDriveRodBox2, ";
                    command.CommandText += "polozenie3=@heightOfDriveRodBox3, szerpuszk3=@thicknessOfDriveRodBox3, glebpuszk3 = @depthOfDriveRodBox3, ";
                    command.CommandText += "polozenie4=@heightOfDriveRodBox4, szerpuszk4=@thicknessOfDriveRodBox4, glebpuszk4 = @depthOfDriveRodBox4, ";
                    command.CommandText += "polozenie5=@heightOfDriveRodBox5, szerpuszk5=@thicknessOfDriveRodBox5, glebpuszk5 = @depthOfDriveRodBox5, ";
                    command.CommandText += "polozenie6=@heightOfDriveRodBox6, szerpuszk6=@thicknessOfDriveRodBox6, glebpuszk6 = @depthOfDriveRodBox6 ";
                    command.CommandText += " WHERE NR_ART=@NR_ART  SET @RowCount = @@ROWCOUNT;";

                    command.Parameters["@nr_art"].Value = item.nr_art;
                    command.Parameters["@handleHeight"].Value = item.handleHeight;
                    command.Parameters["@dornmass"].Value = item.dornmass;
                    
                    command.Parameters["@heightOfDriveRodBox1"].Value = item.heightOfDriveRodBox1;
                    command.Parameters["@depthOfDriveRodBox1"].Value = item.depthOfDriveRodBox1;
                    command.Parameters["@thicknessOfDriveRodBox1"].Value = item.thicknessOfDriveRodBox1;
                    command.Parameters["@shiftOfDriveRodBox1"].Value = item.shiftOfDriveRodBox1;

                    command.Parameters["@heightOfDriveRodBox2"].Value = item.heightOfDriveRodBox2;
                    command.Parameters["@depthOfDriveRodBox2"].Value = item.depthOfDriveRodBox2;
                    command.Parameters["@thicknessOfDriveRodBox2"].Value = item.thicknessOfDriveRodBox2;
                    command.Parameters["@shiftOfDriveRodBox2"].Value = item.shiftOfDriveRodBox2;

                    command.Parameters["@heightOfDriveRodBox3"].Value = item.heightOfDriveRodBox3;
                    command.Parameters["@depthOfDriveRodBox3"].Value = item.depthOfDriveRodBox3;
                    command.Parameters["@thicknessOfDriveRodBox3"].Value = item.thicknessOfDriveRodBox3;
                    command.Parameters["@shiftOfDriveRodBox3"].Value = item.shiftOfDriveRodBox3;

                    command.Parameters["@heightOfDriveRodBox4"].Value = item.heightOfDriveRodBox4;
                    command.Parameters["@depthOfDriveRodBox4"].Value = item.depthOfDriveRodBox4;
                    command.Parameters["@thicknessOfDriveRodBox4"].Value = item.thicknessOfDriveRodBox4;
                    command.Parameters["@shiftOfDriveRodBox4"].Value = item.shiftOfDriveRodBox4;

                    command.Parameters["@heightOfDriveRodBox5"].Value = item.heightOfDriveRodBox5;
                    command.Parameters["@depthOfDriveRodBox5"].Value = item.depthOfDriveRodBox5;
                    command.Parameters["@thicknessOfDriveRodBox5"].Value = item.thicknessOfDriveRodBox5;
                    command.Parameters["@shiftOfDriveRodBox5"].Value = item.shiftOfDriveRodBox5;

                    command.Parameters["@heightOfDriveRodBox6"].Value = item.heightOfDriveRodBox6;
                    command.Parameters["@depthOfDriveRodBox6"].Value = item.depthOfDriveRodBox6;
                    command.Parameters["@thicknessOfDriveRodBox6"].Value = item.thicknessOfDriveRodBox6;
                    command.Parameters["@shiftOfDriveRodBox6"].Value = item.shiftOfDriveRodBox6;

                    //int RowsAffected = Convert.ToInt32(command.Parameters["@RowCount"].Value); dla sprawdzenia czy rekord został odnaleziony w bazie

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
            program.SaveDriveRodsParameters(program._databaseConnectionString);
            //test comment
        }
        

       /* static void Main(string[] args)
        {
            LoadExcelMain(args[0], args[1]);
        }*/
    }

}
