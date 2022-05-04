using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using robot.DataSet1TableAdapters;
using robot.RiskTableAdapters;

namespace robot
{
    class cl_Parser_MKD
    {
        COUNTRY_LogTableAdapter logAdapter;
        string report;

        public void OpenFile() 
        {
            logAdapter = new COUNTRY_LogTableAdapter();
            
            string pathFile = @"C:\Users\Людмила\source\repos\robot\MKD_DCA_03.22.xlsx"; // Путь к файлу отчета
            //static string pathFile = @"C:\Users\Людмила\source\repos\robot\DCA.xlsx"; // Путь к файлу отчета
            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
            Excel.Application ex = new Excel.Application();
            Excel.Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл
            
            if (pathFile.Contains("DCA")) parse_MKD_DCA(ex);
            if (pathFile.Contains("snapshot")) parse_MKD_SNAP(ex);
        }

        public void parse_MKD_DCA(Excel.Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_DCA", "MKD", DateTime.Now, true, report);

            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = sheet.get_Range("A1", last);
            int lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;
            cl_MKD_DCA MKD_DCA = new cl_MKD_DCA();


            int i = 2; // Строка начала периода

            try
            {
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

                    try
                    {
                        ad_MKD_DCA_raw.InsertRow(MKD_DCA.LN, MKD_DCA.Payment_date.ToString("yyyy-MM-dd"), MKD_DCA.DCA_name, MKD_DCA.Payment_amount, MKD_DCA.DCA_comission_amount, MKD_DCA.Reestr_date.ToString("yyyy-MM-dd"));
                        Console.WriteLine((i - 1).ToString() + "/" + (lastUsedRow - 1).ToString() + " row uploaded");
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_DCA", "MKD", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        ex.Quit();
                    }

                    i++;
                }

                SP sp = new SP();
                sp.sp_MKD_TOTAL_DCA(MKD_DCA.Reestr_date);
                Console.WriteLine("Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.");
            }
            catch (Exception exc)
            {
                //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_DCA", "MKD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                ex.Quit();
                return;
            }

            ex.Quit();

            //SP sp = new SP();
            //sp.sp_MKD_TOTAL_DCA(MKD_DCA.Reestr_date);

            report = "Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.";
            logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_DCA", "MKD", DateTime.Now, true, report);

            Console.WriteLine("Do you want to transport DCA to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                TransportDCAToRisk(MKD_DCA.Reestr_date);
            }

            //Console.ReadKey();

            //Xml                                                           ----TO_DO

        }

        private void TransportDCAToRisk(DateTime t_date)
        {
            try
            {
                SPRisk sprisk = new SPRisk();
                sprisk.sp_MKD_TOTAL_DCA(t_date);
                Console.WriteLine("DCA was transported to [Risk].[dbo].[TOTAL_DCA]");
                report = "DCA was transported to [Risk].[dbo].[TOTAL_DCA]";
                logAdapter.InsertRow("cl_Parser_MKD", "TransportDCAToRisk", "MKD", DateTime.Now, true, report);
                //report into log
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MKD", "TransportDCAToRisk", "MKD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }

            Console.ReadKey();

        }

        public void parse_MKD_SNAP(Excel.Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_SNAP", "MKD", DateTime.Now, true, report);

            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = sheet.get_Range("A1", last);
            int lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;
            cl_MKD_SNAP MKD_SNAP = new cl_MKD_SNAP();

            int i = 2; // Строка начала периода

            try
            {
                string fileName = ex.Workbooks.Item[1].Name;
                fileName = fileName.Substring(fileName.IndexOf("_") + 1, 10); //.ToString("yyyy-MM-dd");

                DateTime reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Excel.Range).Value;
                MKD_SNAP.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);

                MKD_SNAP_rawTableAdapter ad_MKD_SNAP_raw = new MKD_SNAP_rawTableAdapter();
                ad_MKD_SNAP_raw.DeletePeriod(MKD_SNAP.Reestr_date.ToString("yyyy-MM-dd"));

                while (i <= lastUsedRow)
                {
                    MKD_SNAP.Loan = (sheet.Cells[i, 1] as Excel.Range).Value.ToString();
                    MKD_SNAP.Current_status = (sheet.Cells[i, 2] as Excel.Range).Value;
                    MKD_SNAP.Loan_disbursement_date = DateTime.Parse((sheet.Cells[i, 3] as Excel.Range).Value);
                    MKD_SNAP.Product = (sheet.Cells[i, 4] as Excel.Range).Value;
                    MKD_SNAP.DPD = (int)(sheet.Cells[i, 5] as Excel.Range).Value;
                    MKD_SNAP.Historical_loan_status = (sheet.Cells[i, 6] as Excel.Range).Value;
                    MKD_SNAP.Principal_balance = (double)(sheet.Cells[i, 7] as Excel.Range).Value;
                    MKD_SNAP.Monthly_fee_balance = (double)(sheet.Cells[i, 8] as Excel.Range).Value;
                    MKD_SNAP.Guarantor_fee_balance = (double)(sheet.Cells[i, 9] as Excel.Range).Value;
                    MKD_SNAP.Penalty_fee_balance = (double)(sheet.Cells[i, 10] as Excel.Range).Value;
                    MKD_SNAP.Penalty_interest_balance = (double)(sheet.Cells[i, 11] as Excel.Range).Value;
                    MKD_SNAP.Interest_balance = (double)(sheet.Cells[i, 12] as Excel.Range).Value;

                    try
                    {
                        ad_MKD_SNAP_raw.InsertRow(MKD_SNAP.Reestr_date.ToString("yyyy-MM-dd"), MKD_SNAP.Loan, MKD_SNAP.Current_status, MKD_SNAP.Loan_disbursement_date.ToString("yyyy-MM-dd"),
                            MKD_SNAP.Product, MKD_SNAP.DPD, MKD_SNAP.Historical_loan_status, MKD_SNAP.Principal_balance, MKD_SNAP.Monthly_fee_balance, MKD_SNAP.Guarantor_fee_balance, MKD_SNAP.Penalty_fee_balance,
                            MKD_SNAP.Penalty_interest_balance, MKD_SNAP.Interest_balance);
                        Console.WriteLine((i - 1).ToString() + "/" + (lastUsedRow - 1).ToString() + " row uploaded");
                    }
                    catch (Exception exc)
                    {
                        COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                        logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_SNAP", "MKD", DateTime.Now,false,exc.Message);
                        Console.WriteLine("Error");
                        ex.Quit();
                    }

                    i++;
                }

                SP sp = new SP();
                sp.sp_MKD2_portfolio_snapshot();
                sp.sp_MKD_TOTAL_SNAP();

                Console.WriteLine("Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.");
            }
            catch (Exception exc)
            {
                COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_SNAP", "MKD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                ex.Quit();
                return;
            }


            ex.Quit();

            report = "Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.";
            logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_SNAP", "MKD", DateTime.Now, true, report);

            Console.WriteLine("Do you want to transport Snap to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                TransportSnapToRisk(MKD_SNAP.Reestr_date);
            }

            //Console.ReadKey();

            //Xml                                                           ----TO_DO

        }

        private void TransportSnapToRisk(DateTime snapdate)
        {
            try
            {
                SPRisk sprisk = new SPRisk();
                sprisk.sp_MKD_TOTAL_SNAP(snapdate);
                Console.WriteLine("Snap was transported to [Risk].[dbo].[MKD2_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP].");
                report = "Snap was transported to [Risk].[dbo].[MKD2_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP].";
                logAdapter.InsertRow("cl_Parser_MKD", "TransportSnapToRisk", "MKD", DateTime.Now, true, report);

                //report
                sprisk.sp_MKD_TOTAL_SNAP_CFIELD();
                Console.WriteLine("[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.");
                report = "[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                logAdapter.InsertRow("cl_Parser_MKD", "TransportSnapToRisk", "MKD", DateTime.Now, true, report);

                //report into log
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MKD", "TransportSnapToRisk", "MKD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }

            Console.ReadKey();

        }

    }
}