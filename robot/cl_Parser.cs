using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using robot.DataSet1TableAdapters;

namespace robot
{
    class cl_Parser
    {


        static string pathFile = @"C:\Users\Людмила\source\repos\robot\DCA.xlsx"; // Путь к файлу отчета
        string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав


        public void parse_MKD_DCA()
        {
            Excel.Application ex = new Excel.Application();
            Excel.Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл


            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = sheet.get_Range("A1", last);
            int lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;


            int i = 2; // Строка начала периода

            cl_MKD_DCA MKD_DCA = new cl_MKD_DCA();
            DateTime reestr_date = (DateTime)(sheet.Cells[i, 2] as Excel.Range).Value;
            MKD_DCA.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);

            MKD_DCA_rawTableAdapter ad_MKD_DCA_raw = new MKD_DCA_rawTableAdapter();
            ad_MKD_DCA_raw.DeletePeriod(MKD_DCA.Reestr_date.ToString("yyyy-MM-dd"));

            while (i <= lastUsedRow)
            {
                MKD_DCA.LN = (int)(sheet.Cells[i, 1] as Excel.Range).Value;
                MKD_DCA.Payment_date = (sheet.Cells[i, 2] as Excel.Range).Value;
                MKD_DCA.DCA_name = (sheet.Cells[i, 3] as Excel.Range).Value;
                MKD_DCA.Payment_amount = (double)(sheet.Cells[i, 4] as Excel.Range).Value;
                MKD_DCA.DCA_comission_amount = (double)(sheet.Cells[i, 5] as Excel.Range).Value;

                ad_MKD_DCA_raw.InsertRow(MKD_DCA.LN, MKD_DCA.Payment_date.ToString("yyyy-MM-dd"), MKD_DCA.DCA_name, MKD_DCA.Payment_amount, MKD_DCA.DCA_comission_amount, MKD_DCA.Reestr_date.ToString("yyyy-MM-dd"));

                // Строка для отправки в БД
                //string sqlinsert = "insert into dwh_risk.dbo.[MKD_DCA_raw] (LN, Payment_date, DCA_name, Payment_Amount, DCA_Comission_amount, reestr_date) " +
                //"values (" + Convert.ToString(MKD_DCA.LN) + ", " + "'" + Convert.ToDateTime(MKD_DCA.Payment_date).ToString("yyyy-MM-dd") + "'" + ", " + "'" 
                //+ MKD_DCA.DCA_name + "'" + ", " + Convert.ToString(MKD_DCA.Payment_amount).Replace(",", ".") + ", "
                //+ Convert.ToString(MKD_DCA.DCA_comission_amount).Replace(",", ".") + ", '" + MKD_DCA.Reestr_date.ToString("yyyy-MM-dd") + "')";

                //Console.WriteLine(sqlinsert);
                //Console.ReadLine();

                //// Подключение к БД
                //SqlConnection objConnect = new SqlConnection();
                //objConnect.ConnectionString = connectionString;
                //objConnect.Open();

                //// Отправвка команды 
                //if (objConnect.State == ConnectionState.Open)
                //{
                //    SqlCommand objCommand = new SqlCommand(sqlinsert);
                //    objCommand.Connection = objConnect;

                //    objCommand.ExecuteNonQuery();
                //}

                i++;
            }

            ex.Quit();

            SP sp = new SP();
            sp.sp_MKD_TOTAL_DCA(MKD_DCA.Reestr_date);

            Console.WriteLine("Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.");
            Console.ReadKey();

            //Xml                                                           ----TO_DO

        }

    }
}