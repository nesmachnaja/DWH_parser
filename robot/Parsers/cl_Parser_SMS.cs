using Microsoft.Office.Interop.Excel;
using robot.DataSet1TableAdapters;
using robot.RiskTableAdapters;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static robot.DataSet1;

namespace robot.Parsers
{
    class cl_Parser_SMS
    {
        private int lastUsedRow;
        COUNTRY_LogTableAdapter logAdapter;
        SP sp = new SP();
        SPRisk sprisk = new SPRisk();
        string report;
        string pathFile;
        string brand = "";
        DateTime reestr_date;
        int success = 0;

        public void StartParsing()
        {
            logAdapter = new COUNTRY_LogTableAdapter();
            int correctPath = 0;

            while (correctPath == 0)
            {
                try
                {
                    pathFile = GetPath();
                    OpenFile(pathFile);
                    correctPath = 1;
                }
                catch
                {
                    Console.WriteLine("Incorrect file path.");
                }
            }
        }

        private static string GetPath()
        {
            Console.WriteLine("Appoint file path: ");
            string pathFile = Console.ReadLine();
            return pathFile;
        }

        public void OpenFile(string pathFile)
        {
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

            SMS_CESS_rawDataTable sms_cess = new SMS_CESS_rawDataTable();
            int i = 2; // Строка начала периода

            try
            {
                string fileName = ex.Workbooks.Item[1].Name;

                if (fileName.Contains("SMS")) brand = "SMSFinance";
                if (fileName.Contains("VIV")) brand = "Vivus";

                fileName = fileName.Replace("ces", "").Replace("prosh", "").Replace("SMS", "").Replace("VIV", "").Replace(".xlsx", "").Insert(2, ".").Insert(5, "."); //.ToString("yyyy-MM-dd");

                reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                //SMS_CESS.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth

                SMS_CESS_rawTableAdapter ad_SMS_CESS_raw = new SMS_CESS_rawTableAdapter();
                ad_SMS_CESS_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"), brand);

                while (i < firstNull)
                {
                    SMS_CESS_rawRow sms_cess_row = sms_cess.NewSMS_CESS_rawRow();

                    sms_cess_row["reestr_date"] = reestr_date;

                    sms_cess_row["Cess_date"] = (DateTime)(sheet.Cells[i, 1] as Range).Value;
                    sms_cess_row["Mobile"] = (sheet.Cells[i, 2] as Range).Value.ToString();
                    sms_cess_row["Loan_id"] = (int)(sheet.Cells[i, 3] as Range).Value;
                    sms_cess_row["Issue_date"] = (DateTime)(sheet.Cells[i, 4] as Range).Value;
                    sms_cess_row["Client_id"] = (int)(sheet.Cells[i, 5] as Range).Value;
                    sms_cess_row["DPD"] = (int)(sheet.Cells[i, 6] as Range).Value;
                    sms_cess_row["OD"] = (double)(sheet.Cells[i, 7] as Range).Value;
                    sms_cess_row["Perc_sroch"] = (double)(sheet.Cells[i, 8] as Range).Value;
                    sms_cess_row["Perc_prosr"] = (double)(sheet.Cells[i, 9] as Range).Value;
                    sms_cess_row["Com_transfer"] = (double)(sheet.Cells[i, 10] as Range).Value;
                    sms_cess_row["Penalty"] = (double)(sheet.Cells[i, 11] as Range).Value;
                    sms_cess_row["Rest_all"] = (double)(sheet.Cells[i, 12] as Range).Value;
                    sms_cess_row["Value"] = (double)(sheet.Cells[i, 13] as Range).Value;
                    sms_cess_row["CC"] = (double)(sheet.Cells[i, 14] as Range).Value;
                    //sms_cess_row["Retdate"] = (DateTime?)(sheet.Cells[i, 15] as Range).Value == null ? (DateTime?)DBNull.Value : (DateTime?)(sheet.Cells[i, 15] as Range).Value;

                    sms_cess_row["brand"] = brand;

                    sms_cess.AddSMS_CESS_rawRow(sms_cess_row);
                    sms_cess.AcceptChanges();

                    Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");

                    i++;
                }

                try
                {
                    sp.sp_SMS_CESS_raw(sms_cess);
                }
                catch (Exception exc)
                {
                    logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, false, exc.Message);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_descr: " + exc.Message);
                    ex.Quit();

                    return;
                }

                //Console.WriteLine("Loading is ready. " + (firstNull - 1).ToString() + " rows were processed.");
            }
            catch (Exception exc)
            {
                //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                ex.Quit();

                return;
            }


            ex.Quit();

            report = "Loading is ready. " + (firstNull - 1).ToString() + " rows were processed.";
            logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_CESS", "SMS", DateTime.Now, true, report);
            Console.WriteLine(report);

            TotalCessForming();

            Console.WriteLine("Do you want to transport snap to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                success = TransportToRisk();
            }

            if (success == 1)
            {
                cl_Send_Report send_report = new cl_Send_Report("SMS_CESS", 1);
                Console.WriteLine("Report was sended.");
            }
        }

        private void TotalCessForming()
        {
            try
            {
                sp.sp_SMS_cession(reestr_date);

                report = "Data was transported to SMS_cession successfully.";
                logAdapter.InsertRow("cl_Parser_SMS", "TotalCessForming", "SMS", DateTime.Now, true, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_SMS", "TotalCessForming", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);

                return;
            }

            try 
            {
                sp.sp_SMS_TOTAL_CESS(reestr_date);

                report = "Data was transported to TOTAL_CESS successfully.";
                logAdapter.InsertRow("cl_Parser_SMS", "TotalCessForming", "SMS", DateTime.Now, true, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_SMS", "TotalCessForming", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);

                return;
            }
        }

        private int TransportToRisk()
        {
            try
            {
                sprisk.sp_SMS_TOTAL_CESS(reestr_date);

                report = "Cessions were transported to their destination on [Risk]";
                logAdapter.InsertRow("cl_Parser_SMS", "TransportToRisk", "SMS", DateTime.Now, true, report);
                Console.WriteLine(report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_SMS", "TransportToRisk", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);

                return 0;
            }

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
            lastUsedRow = last.Row; // Последняя строка в документе

            int i = 2; // Строка начала периода

            int firstNull = SearchFirstNullRow(sheet, lastUsedRow);

            try
            {
                string fileName = ex.Workbooks.Item[1].Name;

                SMS_SNAP_rawDataTable sms_snap = new SMS_SNAP_rawDataTable();

                if (fileName.ToLower().Contains("sms")) brand = "SMSFinance";
                if (fileName.ToLower().Contains("viv")) brand = "Vivus";

                fileName = "01." + fileName.Replace("portf_", "").Replace("smsfin_", "").Replace("vivus_", "").Replace(".xlsx", "").Insert(2, "."); //.Insert(5, "."); //.ToString("yyyy-MM-dd");

                reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                //SMS_SNAP.Reestr_date = reestr_date;       //current date

                SMS_SNAP_rawTableAdapter ad_SMS_SNAP_raw = new SMS_SNAP_rawTableAdapter();
                ad_SMS_SNAP_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"), brand);

                while (i < firstNull)
                {
                    SMS_SNAP_rawRow sms_snap_row = sms_snap.NewSMS_SNAP_rawRow();

                    sms_snap_row["reestr_date"] = reestr_date;

                    sms_snap_row["ID_loan"] = (sheet.Cells[i, 1] as Range).Value.ToString();
                    sms_snap_row["Phone"] = (sheet.Cells[i, 2] as Range).Value.ToString();
                    sms_snap_row["Od"] = (double)(sheet.Cells[i, 3] as Range).Value;
                    sms_snap_row["Com"] = (double)(sheet.Cells[i, 4] as Range).Value;
                    sms_snap_row["Pen_balance"] = (double)(sheet.Cells[i, 5] as Range).Value;
                    sms_snap_row["Od_com"] = (double)(sheet.Cells[i, 6] as Range).Value;
                    sms_snap_row["Day_delay"] = (int)(sheet.Cells[i, 7] as Range).Value;
                    sms_snap_row["Date_start"] = (DateTime)(sheet.Cells[i, 8] as Range).Value;
                    sms_snap_row["ID_client"] = (sheet.Cells[i, 9] as Range).Value.ToString();
                    sms_snap_row["Interest"] = (double)(sheet.Cells[i, 10] as Range).Value;
                    sms_snap_row["Product"] = (sheet.Cells[i, 11] as Range).Value;
                    sms_snap_row["Ces"] = (sheet.Cells[i, 12] as Range).Value;
                    sms_snap_row["Final_interest"] = (double)(sheet.Cells[i, 13] as Range).Value;
                    sms_snap_row["Prod"] = (sheet.Cells[i, 14] as Range).Value;
                    sms_snap_row["Status"] = (sheet.Cells[i, 15] as Range).Value;

                    sms_snap_row["brand"] = brand;

                    sms_snap.AddSMS_SNAP_rawRow(sms_snap_row);
                    sms_snap.AcceptChanges();

                    Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");

                    i++;

                }

                Task task_sms_snap_raw = new Task(() =>
                {
                    sp.sp_SMS_SNAP_raw(sms_snap);
                },
                TaskCreationOptions.LongRunning);

                try
                {
                    task_sms_snap_raw.RunSynchronously();
                    //sp.sp_SMS_SNAP_raw(sms_snap);
                }
                catch (Exception exc)
                {
                    logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_SNAP", "SMS", DateTime.Now, false, exc.Message);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_descr: " + exc.Message);
                    ex.Quit();

                    return;
                }


                report = "Loading is ready. " + (firstNull - 2).ToString() + " rows were processed.";
                logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_SNAP", "SMS", DateTime.Now, true, report);
                Console.WriteLine(report);

            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_SMS", "parse_SMS_SNAP", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                ex.Quit();

                return;
            }


            ex.Quit();

            TotalSnapForming();
            TotalSnapCFForming();

            Console.WriteLine("Do you want to transport Snap to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                TransportSnapToRisk();
                success = TransportSnapCFToRisk();
            }

            if (success == 1)
            {
                cl_Send_Report send_report = new cl_Send_Report("SMS_SNAP", 1);
                Console.WriteLine("Report was sended.");
            }

        }

        private void TotalSnapCFForming()
        {
            try
            {
                sp.sp_SMS_TOTAL_SNAP_CFIELD();

                report = "[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                logAdapter.InsertRow("cl_Parser_SMS", "TotalSnapCFForming", "SMS", DateTime.Now, true, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_SMS", "TotalSnapCFForming", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }
        }

        private void TotalSnapForming()
        {
            try
            {
                sp.sp_SMS_TOTAL_SNAP(reestr_date);

                report = "[dbo].[TOTAL_SNAP] was formed.";
                logAdapter.InsertRow("cl_Parser_SMS", "TotalSnapForming", "SMS", DateTime.Now, true, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_SMS", "TotalSnapForming", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }
        }

        private void TransportSnapToRisk()
        {
            Task task_snap = new Task(() =>
            {
                SPRisk sprisk = new SPRisk();
                sprisk.sp_SMS_TOTAL_SNAP(reestr_date);
            },
            TaskCreationOptions.LongRunning);

            try
            {
                task_snap.RunSynchronously();

                Console.WriteLine("Snap was transported to [Risk].[dbo].[SMS_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP].");
                report = "Snap was transported to [Risk].[dbo].[SMS_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP].";
                logAdapter.InsertRow("cl_Parser_SMS", "TransportSnapToRisk", "SMS", DateTime.Now, true, report);

                //report into log
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_SMS", "TransportSnapToRisk", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }

            //Console.ReadKey();

        }

        private int TransportSnapCFToRisk()
        {
            Task task_snap_cf = new Task(() =>
            {
                SPRisk sprisk = new SPRisk();
                sprisk.sp_SMS_TOTAL_SNAP_CFIELD(reestr_date);
            },
            TaskCreationOptions.LongRunning);

            try
            {
                task_snap_cf.RunSynchronously();

                Console.WriteLine("[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.");
                report = "[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                logAdapter.InsertRow("cl_Parser_SMS", "TransportSnapCFToRisk", "SMS", DateTime.Now, true, report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_SMS", "TransportSnapCFToRisk", "SMS", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }

            //Console.ReadKey();

        }
    }
}
