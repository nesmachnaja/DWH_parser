using Microsoft.Office.Interop.Excel;
using robot.DataSet1TableAdapters;
using robot.RiskTableAdapters;
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
        private int lastUsedRow;
        COUNTRY_LogTableAdapter logAdapter;
        SP sp = new SP();
        SPRisk sprisk = new SPRisk();
        string report;

        public void OpenFile()
        {
            logAdapter = new COUNTRY_LogTableAdapter();

            string pathFile = @"C:\Users\Людмила\source\repos\robot\portf_smsfin_0422.xlsx"; // Путь к файлу отчета
            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
            Application ex = new Application();
            Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл

            if (pathFile.Contains("ces") || pathFile.Contains("prosh")) parse_SMS_CESS(ex);
            if (pathFile.Contains("portf")) parse_SNAP_SNAP(ex);
        }

        private void parse_SMS_CESS(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range range = sheet.get_Range("A1", last);
            int lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;

            int firstNull = SearchFirstNullRow(sheet, lastUsedRow);

            cl_SMS_CESS SMS_CESS = new cl_SMS_CESS();
            int i = 2; // Строка начала периода

            try
            {
                string fileName = ex.Workbooks.Item[1].Name;

                if (fileName.Contains("SMS")) SMS_CESS.Brand = "SMS";
                if (fileName.Contains("VIV")) SMS_CESS.Brand = "Vivus";

                fileName = fileName.Replace("ces", "").Replace("prosh", "").Replace("SMS", "").Replace("VIV", "").Replace(".xlsx", "").Insert(2, ".").Insert(5, "."); //.ToString("yyyy-MM-dd");

                DateTime reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                //SMS_CESS.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                SMS_CESS.Reestr_date = reestr_date;       //current date

                //ex.Quit();

                SMS_CESS_rawTableAdapter ad_SMS_CESS_raw = new SMS_CESS_rawTableAdapter();
                ad_SMS_CESS_raw.DeletePeriod(SMS_CESS.Reestr_date.ToString("yyyy-MM-dd"), SMS_CESS.Brand);

                while (i < firstNull)
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
                            SMS_CESS.Penalty, SMS_CESS.Rest_all, SMS_CESS.Value, SMS_CESS.CC, SMS_CESS.Retdate, SMS_CESS.Brand); //.ToString("yyyy-MM-dd"));
                        Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_descr: " + exc.Message);
                        ex.Quit();
                        Console.ReadKey();
                    }

                    i++;
                }

                SP sp = new SP();
                sp.sp_SMS_cession(SMS_CESS.Reestr_date);
                report = "Data was transported to SMS_cession successfully.";
                logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, true, report);

                sp.sp_SMS_TOTAL_CESS(SMS_CESS.Reestr_date);
                report = "Data was transported to TOTAL_CESS successfully.";
                logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, true, report);

                Console.WriteLine("Loading is ready. " + (firstNull - 1).ToString() + " rows were processed.");
            }
            catch (Exception exc)
            {
                //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                ex.Quit();
                Console.ReadKey();
                return;
            }


            ex.Quit();

            report = "Loading is ready. " + (firstNull - 1).ToString() + " rows were processed.";
            logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, true, report);

            Console.WriteLine("Do you want to transport snap to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                TransportToRisk(SMS_CESS.Reestr_date);
            }

            //report                                                           ----TO_DO

        }

        private void TransportToRisk(DateTime reestr_date)
        {
            try
            {
                reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                SPRisk sprisk = new SPRisk();
                sprisk.sp_SMS_TOTAL_CESS(reestr_date);

                Console.WriteLine("Cessions were transported to their destination on [Risk]");
                report = "Cessions were transported to their destination on [Risk]";
                logAdapter.InsertRow("cl_Parser_SMS", "TransportToRisk", "SMS", DateTime.Now, true, report);

            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_SMS", "TransportToRisk", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
            }

            Console.ReadKey();
        }

        private static int SearchFirstNullRow(Worksheet sheet, int lastUsedRow)
        {
            int firstNull = 0;
            for (int firstEmpty = lastUsedRow + 1; firstEmpty > 1; firstEmpty--)
            {
                if (sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]) != 0 )
                    //&& sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]) == sheet.Application.WorksheetFunction.CountA(sheet.Rows[1]))
                {
                    firstNull = firstEmpty + 1;
                    break;
                }
            }

            return firstNull;
        }

        private void parse_SNAP_SNAP(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_SNAP", "SMS", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range range = sheet.get_Range("A1", last);
            lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;
            cl_SMS_SNAP SMS_SNAP = new cl_SMS_SNAP();

            int i = 2; // Строка начала периода

            int firstNull = SearchFirstNullRow(sheet, lastUsedRow);

            try
            {
                string fileName = ex.Workbooks.Item[1].Name;

                if (fileName.ToLower().Contains("sms")) SMS_SNAP.Brand = "SMS";
                if (fileName.ToLower().Contains("viv")) SMS_SNAP.Brand = "Vivus";

                fileName = "01." + fileName.Replace("portf_", "").Replace("smsfin_", "").Replace("vivus_", "").Replace(".xlsx", "").Insert(2, "."); //.Insert(5, "."); //.ToString("yyyy-MM-dd");

                DateTime reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                SMS_SNAP.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                //SMS_SNAP.Reestr_date = reestr_date;       //current date

                SMS_SNAP_rawTableAdapter ad_SMS_SNAP_raw = new SMS_SNAP_rawTableAdapter();
                ad_SMS_SNAP_raw.DeletePeriod(SMS_SNAP.Reestr_date.ToString("yyyy-MM-dd"), SMS_SNAP.Brand);

                while (i < firstNull)
                {
                    SMS_SNAP.ID_loan = (sheet.Cells[i, 1] as Range).Value.ToString();
                    SMS_SNAP.Phone = (sheet.Cells[i, 2] as Range).Value.ToString();
                    SMS_SNAP.Od = (double)(sheet.Cells[i, 3] as Range).Value;
                    SMS_SNAP.Com = (double)(sheet.Cells[i, 4] as Range).Value;
                    SMS_SNAP.Pen_balance = (double)(sheet.Cells[i, 5] as Range).Value;
                    SMS_SNAP.Od_com = (double)(sheet.Cells[i, 6] as Range).Value;
                    SMS_SNAP.Day_delay = (int)(sheet.Cells[i, 7] as Range).Value;
                    SMS_SNAP.Date_start = (DateTime)(sheet.Cells[i, 8] as Range).Value;
                    SMS_SNAP.ID_client = (sheet.Cells[i, 9] as Range).Value.ToString();
                    SMS_SNAP.Interest = (double)(sheet.Cells[i, 10] as Range).Value;
                    SMS_SNAP.Product = (sheet.Cells[i, 11] as Range).Value;
                    SMS_SNAP.Ces = (sheet.Cells[i, 12] as Range).Value;
                    SMS_SNAP.Final_interest = (double)(sheet.Cells[i, 13] as Range).Value;
                    SMS_SNAP.Prod = (sheet.Cells[i, 14] as Range).Value;
                    SMS_SNAP.Status = (sheet.Cells[i, 15] as Range).Value;

                    try
                    {
                        ad_SMS_SNAP_raw.InsertRow(SMS_SNAP.Reestr_date.ToString("yyyy-MM-dd"), SMS_SNAP.ID_loan, SMS_SNAP.Phone, SMS_SNAP.Od, SMS_SNAP.Com,
                            SMS_SNAP.Pen_balance, SMS_SNAP.Od_com, SMS_SNAP.Day_delay, SMS_SNAP.Date_start.ToString("yyyy-MM-dd"), SMS_SNAP.ID_client, SMS_SNAP.Interest,
                            SMS_SNAP.Product, SMS_SNAP.Ces, SMS_SNAP.Final_interest, SMS_SNAP.Prod, SMS_SNAP.Status, SMS_SNAP.Brand);
                        Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_SNAP", "SMS", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_descr: " + exc.Message);
                        ex.Quit();
                    }

                    i++;
                }


                Console.WriteLine("Loading is ready. " + (firstNull - 2).ToString() + " rows were processed.");
                report = "Loading is ready. " + (firstNull - 2).ToString() + " rows were processed.";
                logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_SNAP", "SMS", DateTime.Now, true, report);

            }
            catch (Exception exc)
            {
                //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_SNAP", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                ex.Quit();
            }


            ex.Quit();

            Console.WriteLine("Do you want to transport Snap to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                TransportSnapToRisk(SMS_SNAP.Reestr_date);
            }

        }

        private void TransportSnapToRisk(DateTime snapdate)
        {
            try
            {
                SPRisk sprisk = new SPRisk();
                sprisk.sp_SMS_TOTAL_SNAP(snapdate);
                Console.WriteLine("Snap was transported to [Risk].[dbo].[SMS_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP].");
                report = "Snap was transported to [Risk].[dbo].[SMS_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP].";
                logAdapter.InsertRow("cl_Parser_SMS", "TransportSnapToRisk", "SMS", DateTime.Now, true, report);

                //report
                sprisk.sp_SMS_TOTAL_SNAP_CFIELD();
                Console.WriteLine("[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.");
                report = "[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                logAdapter.InsertRow("cl_Parser_SMS", "TransportSnapToRisk", "SMS", DateTime.Now, true, report);

                //report into log
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_SMS", "TransportSnapToRisk", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }

            Console.ReadKey();

        }
    }
}
