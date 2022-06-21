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
    class cl_Parser_LIGA
    {
        private int lastUsedRow;
        COUNTRY_LogTableAdapter logAdapter;
        SP sp = new SP();
        SPRisk sprisk = new SPRisk();
        string report;

        public void OpenFile()
        {
            logAdapter = new COUNTRY_LogTableAdapter();

            Console.WriteLine("Appoint file path: ");
            string pathFile = Console.ReadLine();

            //string pathFile = @"C:\Users\Людмила\source\repos\robot\Портфель ЛД 05.2022.xlsb"; // Путь к файлу отчета
            //static string pathFile = @"C:\Users\Людмила\source\repos\robot\DCA.xlsx"; // Путь к файлу отчета
            
            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
            Application ex = new Application();
            Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл

            //if (pathFile.Contains("external_collection")) parse_LIGA_DCA(ex);
            if (pathFile.ToLower().Contains("портфель")) parse_LIGA_SNAP(ex);
        }

        private void parse_LIGA_SNAP(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range range = sheet.get_Range("A1", last);
            lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;
            cl_LIGA_SNAP LIGA_SNAP = new cl_LIGA_SNAP();

            int i = 2; // Строка начала периода
            int firstNull = 0;

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

            try
            {
                string fileName = ex.Workbooks.Item[1].Name;
                fileName = fileName.Replace(".xlsb", "").Replace("xlsx", "").Replace("Портфель ЛД ", ""); //+ DateTime.Now.Year.ToString(); //.ToString("yyyy-MM-dd");

                DateTime reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                LIGA_SNAP.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                //LIGA_SNAP.Reestr_date = reestr_date;       //current date

                LIGA_SNAP_rawTableAdapter ad_LIGA_SNAP_raw = new LIGA_SNAP_rawTableAdapter();
                ad_LIGA_SNAP_raw.DeletePeriod(LIGA_SNAP.Reestr_date.ToString("yyyy-MM-dd"));

                while (i < firstNull)
                {
                    LIGA_SNAP.Disbursement_date = (DateTime)(sheet.Cells[i, 2] as Range).Value;
                    LIGA_SNAP.Loan_id = (sheet.Cells[i, 3] as Range).Value.ToString();
                    LIGA_SNAP.Client_id = (sheet.Cells[i, 4] as Range).Value.ToString();
                    LIGA_SNAP.Loan_amount = (double)(sheet.Cells[i, 6] as Range).Value;
                    LIGA_SNAP.Interest_rate = (double)(sheet.Cells[i, 7] as Range).Value;
                    LIGA_SNAP.Product_raw = (sheet.Cells[i, 9] as Range).Value;
                    LIGA_SNAP.Client_cycle = (double)(sheet.Cells[i, 12] as Range).Value;
                    LIGA_SNAP.Principal = (double)(sheet.Cells[i, 13] as Range).Value;
                    LIGA_SNAP.Interest = (double)(sheet.Cells[i, 14] as Range).Value;
                    LIGA_SNAP.Overdue_principal = (double)(sheet.Cells[i, 15] as Range).Value;
                    LIGA_SNAP.Overdue_interest = (double)(sheet.Cells[i, 16] as Range).Value;
                    LIGA_SNAP.DPD = (int)(sheet.Cells[i, 17] as Range).Value;
                    LIGA_SNAP.Prepayment = (double)(sheet.Cells[i, 18] as Range).Value;
                    LIGA_SNAP.Status = (sheet.Cells[i, 19] as Range).Value;

                    try
                    {
                        ad_LIGA_SNAP_raw.InsertRow(LIGA_SNAP.Reestr_date.ToString("yyyy-MM-dd"), LIGA_SNAP.Disbursement_date.ToString("yyyy-MM-dd"), LIGA_SNAP.Loan_id, LIGA_SNAP.Client_id, LIGA_SNAP.Loan_amount,
                            LIGA_SNAP.Interest_rate, LIGA_SNAP.Product_raw, LIGA_SNAP.Client_cycle, LIGA_SNAP.Principal, LIGA_SNAP.Interest, LIGA_SNAP.Overdue_principal,
                            LIGA_SNAP.Overdue_interest, LIGA_SNAP.DPD, LIGA_SNAP.Prepayment, LIGA_SNAP.Status);
                        Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_descr: " + exc.Message);
                        ex.Quit();
                    }

                    i++;
                }


                Console.WriteLine("Loading is ready. " + (firstNull - 2).ToString() + " rows were processed.");
                report = "Loading is ready. " + (firstNull - 2).ToString() + " rows were processed.";
                logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, true, report);
                
                sp.sp_LIGA_portfolio_snapshot(LIGA_SNAP.Reestr_date);
                sp.sp_LIGA_TOTAL_SNAP(LIGA_SNAP.Reestr_date);

                report = "[LIGA_portfolio_snapshot] and [TOTAL_SNAP] were formed.";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, true, report);

                //sp.sp_LIGA_TOTAL_SNAP_CFIELD();
                //Console.WriteLine("[DWH_Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.");
                //report = "[DWH_Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                //logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, true, report);

            }
            catch (Exception exc)
            {
                //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_LIGA", "parse_LIGA_SNAP", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                ex.Quit();
            }


            ex.Quit();

            Console.WriteLine("Do you want to transport Snap to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                TransportSnapToRisk(LIGA_SNAP.Reestr_date);
            }

            //Console.ReadKey();

            //report                                                           ----TO_DO
        }

        private void TransportSnapToRisk(DateTime snapdate)
        {
            try
            {
                SPRisk sprisk = new SPRisk();
                sprisk.sp_LIGA_TOTAL_SNAP(snapdate);
                Console.WriteLine("Snap was transported to [Risk].[dbo].[LIGA_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP].");
                report = "Snap was transported to [Risk].[dbo].[LIGA_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP].";
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportSnapToRisk", "LIGA", DateTime.Now, true, report);

                //report
                sprisk.sp_LIGA_TOTAL_SNAP_CFIELD();
                Console.WriteLine("[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.");
                report = "[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportSnapToRisk", "LIGA", DateTime.Now, true, report);

                //report into log
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_LIGA", "TransportSnapToRisk", "LIGA", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }

            Console.ReadKey();

        }

    }
}
