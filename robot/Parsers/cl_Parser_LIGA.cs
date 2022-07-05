﻿using Microsoft.Office.Interop.Excel;
using robot.DataSet1TableAdapters;
using robot.RiskTableAdapters;
using robot.Structures;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static robot.DataSet1;

namespace robot.Parsers
{
    class cl_Parser_LIGA
    {
        private int lastUsedRow;
        COUNTRY_LogTableAdapter logAdapter;
        SP sp = new SP();
        SPRisk sprisk = new SPRisk();
        string report;
        string pathFile;

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

            if (pathFile.ToLower().Contains("портфель")) parse_LIGA_SNAP(ex);
        }

        private void parse_LIGA_SNAP(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            lastUsedRow = last.Row; // Последняя строка в документе
            LIGA_SNAP_rawDataTable liga_snap = new LIGA_SNAP_rawDataTable();
            DateTime reestr_date;

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

                LIGA_SNAP_rawTableAdapter ad_LIGA_SNAP_raw = new LIGA_SNAP_rawTableAdapter();
                ad_LIGA_SNAP_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"));

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
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                ex.Quit();

                return;
            }


            ex.Quit();

            TotalSnapForming(reestr_date);
            TotalSnapCFForming();

            Console.WriteLine("Do you want to transport Snap to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                TransportSnapToRisk(reestr_date);
                TransportSnapCFToRisk(reestr_date);
            }

            cl_Send_Report send_report = new cl_Send_Report("LIGA_SNAP", 1);
            Console.WriteLine("Report was sended.");

        }

        private void TotalSnapCFForming()
        {
            try 
            {
                sp.sp_LIGA_TOTAL_SNAP_CFIELD();

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

        private void TotalSnapForming(DateTime reestr_date)
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

        private void TransportSnapToRisk(DateTime snapdate)
        {
            Task task_liga_snap = new Task(() =>
            {
                sprisk.sp_LIGA_TOTAL_SNAP(snapdate);
            },
            TaskCreationOptions.LongRunning);

            try
            {
                task_liga_snap.RunSynchronously();

                report = "Snap was transported to [Risk].[dbo].[LIGA_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP].";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportSnapToRisk", "LIGA", DateTime.Now, true, report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportSnapToRisk", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }
        }

        private void TransportSnapCFToRisk(DateTime snapdate)
        {
            Task task_liga_snap = new Task(() =>
            {
                sprisk.sp_LIGA_TOTAL_SNAP_CFIELD(snapdate);
            },
            TaskCreationOptions.LongRunning);

            try
            {
                task_liga_snap.RunSynchronously();

                report = "Snap_CF was transported to [Risk].[dbo].[TOTAL_SNAP_CFIELD].";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportSnapCFToRisk", "LIGA", DateTime.Now, true, report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportSnapCFToRisk", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }
        }

    }
}
