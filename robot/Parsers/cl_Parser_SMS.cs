using Microsoft.Office.Interop.Excel;
using robot.DataSet1TableAdapters;
using robot.Structures;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace robot.Parsers
{
    class cl_Parser_SMS
    {
        COUNTRY_LogTableAdapter logAdapter;

        public void OpenFile()
        {
            logAdapter = new COUNTRY_LogTableAdapter();

            string pathFile = @"C:\Users\Людмила\source\repos\robot\cesSMS29012022.xlsx"; // Путь к файлу отчета
            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
            Application ex = new Application();
            Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл

            if (pathFile.Contains("ces") || pathFile.Contains("prosh")) parse_SMS_CESS(ex);
            //if (pathFile.Contains("snapshot")) parse_SNAP_SNAP(ex);
        }

        private void parse_SMS_CESS(Application ex)
        {
            string report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range range = sheet.get_Range("A1", last);
            int lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;

            int i = 2; // Строка начала периода

            try
            {

                cl_SMS_CESS SMS_CESS = new cl_SMS_CESS();

                string fileName = ex.Workbooks.Item[1].Name;
                fileName = fileName.Replace("ces","").Replace("prosh","").Replace("SMS","").Replace("VIV","").Replace(".xlsx","").Insert(2,".").Insert(5,"."); //.ToString("yyyy-MM-dd");

                DateTime reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                //SMS_CESS.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                SMS_CESS.Reestr_date = reestr_date;       //current date

                //ex.Quit();

                SMS_CESS_rawTableAdapter ad_SMS_CESS_raw = new SMS_CESS_rawTableAdapter();
                ad_SMS_CESS_raw.DeletePeriod(SMS_CESS.Reestr_date.ToString("yyyy-MM-dd"));

                while (i <= lastUsedRow)
                {
                    SMS_CESS.Cess_date = (DateTime)(sheet.Cells[i, 1] as Range).Value;
                    //SMS_CESS.Cess_date = DateTime.Parse((sheet.Cells[i, 1] as Range).Value);
                    SMS_CESS.Mobile = (sheet.Cells[i, 2] as Range).Value.ToString();
                    SMS_CESS.Loan_id = (int)(sheet.Cells[i, 3] as Range).Value;
                    SMS_CESS.Issue_date = (DateTime)(sheet.Cells[i, 4] as Range).Value;
                    SMS_CESS.Client_id = (int)(sheet.Cells[i, 5] as Range).Value;
                    SMS_CESS.DPD = (int)(sheet.Cells[i, 6] as Range).Value;
                    SMS_CESS.OD = (double)(sheet.Cells[i, 7] as Range).Value;
                    SMS_CESS.Perc_sroch = (double)(sheet.Cells[i, 8] as Range).Value;
                    SMS_CESS.Perc_prosr = (double)(sheet.Cells[i, 9] as Range).Value;
                    SMS_CESS.Com_transfer = (double)(sheet.Cells[i, 10] as Range).Value;
                    SMS_CESS.Penalty = (double)(sheet.Cells[i, 11] as Range).Value;
                    SMS_CESS.Rest_all = (double)(sheet.Cells[i, 12] as Range).Value;
                    SMS_CESS.Value = (double)(sheet.Cells[i, 13] as Range).Value;
                    SMS_CESS.CC = (double)(sheet.Cells[i, 14] as Range).Value;
                    SMS_CESS.Retdate = (DateTime?)(sheet.Cells[i, 15] as Range).Value;

                    try
                    {
                        ad_SMS_CESS_raw.InsertRow(SMS_CESS.Reestr_date.ToString("yyyy-MM-dd"), SMS_CESS.Cess_date.ToString("yyyy-MM-dd"), SMS_CESS.Mobile, SMS_CESS.Loan_id, SMS_CESS.Issue_date.ToString("yyyy-MM-dd"),
                            SMS_CESS.Client_id, SMS_CESS.DPD, SMS_CESS.OD, SMS_CESS.Perc_sroch, SMS_CESS.Perc_prosr, SMS_CESS.Com_transfer,
                            SMS_CESS.Penalty, SMS_CESS.Rest_all, SMS_CESS.Value, SMS_CESS.CC, SMS_CESS.Retdate); //.ToString("yyyy-MM-dd"));
                        Console.WriteLine((i - 1).ToString() + "/" + (lastUsedRow - 1).ToString() + " row uploaded");
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        ex.Quit();
                    }

                    i++;
                }

                SP sp = new SP();
                sp.sp_SMS_cession(SMS_CESS.Reestr_date);
                sp.sp_SMS_TOTAL_CESS(SMS_CESS.Reestr_date);

                Console.WriteLine("Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.");
            }
            catch (Exception exc)
            {
                //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                ex.Quit();
                return;
            }


            ex.Quit();

            report = "Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.";
            logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, true, report);

            Console.ReadKey();

            //report                                                           ----TO_DO

        }
    }
}
