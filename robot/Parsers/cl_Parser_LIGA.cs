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
    class cl_Parser_LIGA : cl_Parser
    {
        //private int lastUsedRow;
        //COUNTRY_LogTableAdapter logAdapter;
        //SP sp = new SP();
        //SPRisk sprisk = new SPRisk();
        //DateTime reestr_date;
        //string report;
        //string pathFile;
        //int success = 0;

        LIGA_CESS_rawDataTable LIGA_cess = new LIGA_CESS_rawDataTable();
        string file_type = "";
        DateTime cession_date;

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

            if (pathFile.ToLower().Contains("портфель") 
                && !pathFile.ToLower().Contains("банкроты") && !pathFile.ToLower().Contains("цессии")) parse_LIGA_SNAP(ex);
            if (pathFile.ToLower().Contains("банкроты") || pathFile.ToLower().Contains("цессии")) parse_LIGA_CESS(ex);
        }

        private void parse_LIGA_SNAP(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            lastUsedRow = last.Row; // Последняя строка в документе
            LIGA_SNAP_rawDataTable liga_snap = new LIGA_SNAP_rawDataTable();

            int i = 2; // Строка начала периода
            int firstNull = 0;

            firstNull = SearchFirstNullRow(sheet, firstNull);

            try
            {
                string fileName = ex.Workbooks.Item[1].Name;
                fileName = "01." + fileName.Replace(".xlsb", "").Replace(".xlsx", "").Replace("Портфель ЛД ", "").Replace("Портфель_ЛД_", ""); //+ DateTime.Now.Year.ToString(); //.ToString("yyyy-MM-dd");

                reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                //LIGA_SNAP.Reestr_date = reestr_date;       //current date

                while (i < firstNull)
                {
                    LIGA_SNAP_rawRow liga_snap_raw = liga_snap.NewLIGA_SNAP_rawRow();

                    liga_snap_raw["Reestr_date"] = reestr_date;

                    liga_snap_raw["Disbursement_date"] = (DateTime)(sheet.Cells[i, 2] as Range).Value;
                    liga_snap_raw["Loan_id"] = (sheet.Cells[i, 3] as Range).Value.ToString();
                    liga_snap_raw["Client_id"] = (sheet.Cells[i, 4] as Range).Value.ToString();
                    liga_snap_raw["Loan_amount"] = (double)(sheet.Cells[i, 6] as Range).Value;
                    liga_snap_raw["Interest_rate"] = (double)(sheet.Cells[i, 7] as Range).Value;
                    liga_snap_raw["Product_raw"] = (sheet.Cells[i, 9] as Range).Value;
                    liga_snap_raw["Client_cycle"] = (double)(sheet.Cells[i, 12] as Range).Value;
                    liga_snap_raw["Principal"] = (double)(sheet.Cells[i, 13] as Range).Value;
                    liga_snap_raw["Interest"] = (double)(sheet.Cells[i, 14] as Range).Value;
                    liga_snap_raw["Overdue_principal"] = (double)(sheet.Cells[i, 15] as Range).Value;
                    liga_snap_raw["Overdue_interest"] = (double)(sheet.Cells[i, 16] as Range).Value;
                    liga_snap_raw["DPD"] = (int)(sheet.Cells[i, 17] as Range).Value;
                    liga_snap_raw["Prepayment"] = (double)(sheet.Cells[i, 18] as Range).Value;
                    liga_snap_raw["Status"] = (sheet.Cells[i, 19] as Range).Value;

                    liga_snap.AddLIGA_SNAP_rawRow(liga_snap_raw);
                    liga_snap.AcceptChanges();

                    Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");

                    i++;
                }

                if (liga_snap.Rows.Count > 0)
                {
                    LIGA_SNAP_rawTableAdapter ad_LIGA_SNAP_raw = new LIGA_SNAP_rawTableAdapter();
                    ad_LIGA_SNAP_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"));

                    try
                    {
                        sp.sp_LIGA_SNAP_raw(liga_snap);
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_descr: " + exc.Message);
                        ex.Quit();

                        return;
                    }

                    report = "Loading is ready. " + (firstNull - 2).ToString() + " rows were processed.";
                    Console.WriteLine(report);
                    logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, true, report);
                }
                else
                {
                    report = "File was empty. There is no one row.";
                    logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, false, report);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_descr: " + report);
                    ex.Quit();

                    return;
                }

            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, false, exc.Message);
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
            else
            {
                Console.WriteLine("Do you want to continue? Y - Yes, N - No");
                reply = Console.ReadKey().Key.ToString();

                if (reply.Equals("Y"))
                {
                    StartParsing();
                    return;
                }
            }

            if (success == 1)
            {
                cl_Send_Report send_report = new cl_Send_Report("LIGA_SNAP", 1);
                //Console.WriteLine("Report was sended.");
            }
        }

        private void TotalSnapCFForming()
        {
            try 
            {
                cl_Tasks task = new cl_Tasks("exec DWH_Risk.dbo.sp_LIGA_TOTAL_SNAP_CFIELD");
                //sp.sp_LIGA_TOTAL_SNAP_CFIELD();

                report = "[DWH_Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_LIGA", "TotalSnapCFForming", "LIGA", DateTime.Now, true, report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "TotalSnapCFForming", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);

                return;
            }
        }

        private void TotalSnapForming()
        {
            try
            {
                sp.sp_LIGA_portfolio_snapshot(reestr_date);
                sp.sp_LIGA_TOTAL_SNAP(reestr_date);

                report = "[LIGA_portfolio_snapshot] and [TOTAL_SNAP] were formed.";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_LIGA", "TotalSnapForming", "LIGA", DateTime.Now, true, report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "TotalSnapForming", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);

                return;
            }
        }

        private int SearchFirstNullRow(Worksheet sheet, int firstNull)
        {
            for (int firstEmpty = lastUsedRow; firstEmpty > 1; firstEmpty--)
            {
                if (sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]) != 0 &&
                        sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]) == sheet.Application.WorksheetFunction.CountA(sheet.Rows[1]))
                {
                    //string a = sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]).ToString();
                    //string w = sheet.Application.WorksheetFunction.CountA(sheet.Rows[1]).ToString();
                    firstNull = firstEmpty + 1;
                    break;
                }
            }

            return firstNull;
        }

        private int TransportSnapToRisk()
        {
            Task task_liga_snap = new Task(() =>
            {
                sprisk.sp_LIGA_TOTAL_SNAP(reestr_date);
            },
            TaskCreationOptions.LongRunning);

            try
            {
                task_liga_snap.RunSynchronously();

                report = "Snap was transported to [Risk].[dbo].[LIGA_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP].";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportSnapToRisk", "LIGA", DateTime.Now, true, report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportSnapToRisk", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }
        }

        private int TransportSnapCFToRisk()
        {
            //Task task_liga_snap = new Task(() =>
            //{
            //    sprisk.sp_LIGA_TOTAL_SNAP_CFIELD(reestr_date);
            //},
            //TaskCreationOptions.LongRunning);

            try
            {
                cl_Tasks task = new cl_Tasks("exec Risk.dbo.sp_LIGA_TOTAL_SNAP_CFIELD @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //task_liga_snap.RunSynchronously();

                report = "Snap_CF was transported to [Risk].[dbo].[TOTAL_SNAP_CFIELD].";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportSnapCFToRisk", "LIGA", DateTime.Now, true, report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportSnapCFToRisk", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }
        }


        private void parse_LIGA_CESS(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_CESS", "LIGA", DateTime.Now, true, report);


            string fileName = ex.Workbooks.Item[1].Name;
            int list_count = ex.Worksheets.Count;

            fileName = fileName.Contains("Банкроты") ? fileName.Replace("Банкроты Лига Денег_", "").Substring(0, 10) : fileName.Substring(fileName.IndexOf("от ") + 3, 10); //.ToString("yyyy-MM-dd");
            file_type = ex.Workbooks.Item[1].Name.Contains("Банкроты") ? "Bankrupts" : "Cessions"; //.ToString("yyyy-MM-dd");

            reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
            cession_date = reestr_date;
            reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
            
            int row_count = 0;

            if (file_type == "Bankrupts")
            {
                for (int j = 1; j <= list_count; j++)
                {
                    Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(j);
                    Console.WriteLine("Sheet #" + j.ToString());
                    row_count += parse_LIGA_CESS_bankrupts(sheet);
                }
            }
            else if (file_type == "Cessions")
            {
                Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
                row_count += parse_LIGA_CESS_cessions(sheet);
            }

            if (LIGA_cess.Rows.Count > 0)
            {
                LIGA_CESS_rawTableAdapter ad_LIGA_CESS_raw = new LIGA_CESS_rawTableAdapter();
                ad_LIGA_CESS_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"),file_type);

                try
                {
                    sp.sp_LIGA_CESS_raw(LIGA_cess);
                }
                catch (Exception exc)
                {
                    logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_CESS", "LIGA", DateTime.Now, false, exc.Message);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_descr: " + exc.Message);
                    ex.Quit();
                    //Console.ReadKey();

                    return;
                }

                report = "Loading is ready. " + (row_count).ToString() + " rows were processed.";
                logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_CESS", "LIGA", DateTime.Now, true, report);
                Console.WriteLine(report);

                //LIGA2_cessions_forming(reestr_date);
                LIGA_Total_CESS_forming(reestr_date);
                TotalSnapCFForming();

                Console.WriteLine("Do you want to transport snap to Risk? Y - Yes, N - No");
                string reply = Console.ReadKey().Key.ToString();


                if (reply.Equals("Y"))
                {
                    success = TransportToRisk();
                    success += TransportSnapCFToRisk();
                }
                else
                {
                    Console.WriteLine("Do you want to continue? Y - Yes, N - No");
                    reply = Console.ReadKey().Key.ToString();

                    if (reply.Equals("Y"))
                    {
                        StartParsing();
                        return;
                    }
                }

                if (success == 2)
                {
                    cl_Send_Report send_report = new cl_Send_Report("LIGA_CESS", 1);
                    //Console.WriteLine("Report was sended.");
                }
            }
            else
            {
                report = "File was empty. There is no one row.";
                logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_CESS", "LIGA", DateTime.Now, false, report);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + report);
                ex.Quit();
                Console.ReadKey();

                return;
            }

            ex.Quit();

        }


        private int parse_LIGA_CESS_bankrupts(Worksheet sheet)
        {
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastUsedRow = last.Row; // Последняя строка в документе

            int firstNull = SearchFirstNullRow(sheet, lastUsedRow);

            int i = 2; // Строка начала периода

            try
            {
                while (i < firstNull)
                {
                    LIGA_CESS_rawRow row = LIGA_cess.NewLIGA_CESS_rawRow();

                    row["Reestr_date"] = reestr_date;     //eomonth

                    row["a_id"] = (sheet.Cells[i, 1] as Range).Value;
                    row["cess_date"] = (DateTime)(sheet.Cells[i, 2] as Range).Value;
                    row["contract_id"] = (sheet.Cells[i, 3] as Range).Value;
                    row["client_id"] = (sheet.Cells[i, 4] as Range).Value;
                    row["loan_date"] = (DateTime)(sheet.Cells[i, 5] as Range).Value;
                    row["loan_amount"] = (double)(sheet.Cells[i, 11] as Range).Value;
                    row["rate"] = (double)(sheet.Cells[i, 12] as Range).Value;
                    row["product"] = (sheet.Cells[i, 10] as Range).Value;
                    row["client_cycle"] = (double)(sheet.Cells[i, 13] as Range).Value;
                    row["principal"] = (double)(sheet.Cells[i, 14] as Range).Value;
                    row["interest"] = (double)(sheet.Cells[i, 15] as Range).Value;
                    row["DPD"] = (int)(sheet.Cells[i, 18] as Range).Value;
                    row["status"] = (sheet.Cells[i, 20] as Range).Value;
                    //row["last_payment_date"] = null; //(int)(sheet.Cells[i, 11] as Range).Value;
                    row["last_payment_amount"] = 0; //(int)(sheet.Cells[i, 11] as Range).Value;
                    row["sum_payments"] = 0; //(int)(sheet.Cells[i, 11] as Range).Value;
                    row["recovery_amount"] = 0; //(int)(sheet.Cells[i, 11] as Range).Value;
                    row["file_type"] = file_type; //(int)(sheet.Cells[i, 11] as Range).Value;

                    LIGA_cess.AddLIGA_CESS_rawRow(row);
                    LIGA_cess.AcceptChanges();

                    Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");

                    i++;
                }
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_CESS", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                sheet.Application.Quit();
                Console.ReadKey();

                return 0;
            }

            return firstNull - 2;
        }
        
        private int parse_LIGA_CESS_cessions(Worksheet sheet)
        {
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastUsedRow = last.Row; // Последняя строка в документе

            int firstNull = SearchFirstNullRow(sheet, lastUsedRow);

            int i = 2; // Строка начала периода

            try
            {
                while (i < firstNull)
                {
                    LIGA_CESS_rawRow row = LIGA_cess.NewLIGA_CESS_rawRow();

                    row["Reestr_date"] = reestr_date;     //eomonth

                    row["a_id"] = (sheet.Cells[i, 1] as Range).Value;
                    row["cess_date"] = cession_date;
                    row["contract_id"] = (sheet.Cells[i, 3] as Range).Value;
                    row["client_id"] = (sheet.Cells[i, 2] as Range).Value;
                    row["loan_date"] = (DateTime)(sheet.Cells[i, 16] as Range).Value;
                    row["loan_amount"] = (double)(sheet.Cells[i, 20] as Range).Value;
                    row["rate"] = (double)(sheet.Cells[i, 21] as Range).Value;
                    row["product"] = (sheet.Cells[i, 24] as Range).Value;
                    row["client_cycle"] = (double)(sheet.Cells[i, 25] as Range).Value;
                    row["principal"] = (double)(sheet.Cells[i, 26] as Range).Value;
                    row["interest"] = (double)(sheet.Cells[i, 27] as Range).Value;
                    row["DPD"] = (int)(sheet.Cells[i, 31] as Range).Value;
                    row["status"] = "";
                    row["last_payment_amount"] = 0; //(int)(sheet.Cells[i, 11] as Range).Value;
                    row["sum_payments"] = 0; //(int)(sheet.Cells[i, 11] as Range).Value;
                    row["recovery_amount"] = (int)(sheet.Cells[i, 46] as Range).Value;
                    row["file_type"] = file_type;

                    LIGA_cess.AddLIGA_CESS_rawRow(row);
                    LIGA_cess.AcceptChanges();

                    Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");

                    i++;
                }
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_CESS", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                sheet.Application.Quit();
                Console.ReadKey();

                return 0;
            }

            return firstNull - 2;
        }

        private void LIGA2_cessions_forming(DateTime reestr_date)
        {
            object result;

            Task task_cess = new Task(() =>
            {
                result = sp.sp_LIGA2_cessions(reestr_date);
            },
            TaskCreationOptions.LongRunning);

            try
            {
                task_cess.RunSynchronously();

                report = "LIGA2_cessions was formed successfully.";
                logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_CESS", "LIGA", DateTime.Now, true, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "LIGA_Totoal_CESS_forming", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }
        }
        

        private void LIGA_Total_CESS_forming(DateTime reestr_date)
        {
            object result;

            Task task_cess = new Task(() =>
            {
                result = sp.sp_LIGA_TOTAL_CESS(reestr_date);
            },
            TaskCreationOptions.LongRunning);

            try
            {
                task_cess.RunSynchronously();

                report = "TOTAL_CESS was formed successfully.";
                logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_CESS", "LIGA", DateTime.Now, true, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "LIGA_Totoal_CESS_forming", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }
        }


        private int TransportToRisk()
        {
            try
            {
                sprisk.sp_LIGA_TOTAL_CESS(reestr_date);

                report = "Cessions were transported to their destination on [Risk]";
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportToRisk", "LIGA", DateTime.Now, true, report);
                Console.WriteLine(report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportToRisk", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);

                return 0;
            }

        }
    }
}
