using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using robot.DataSet1TableAdapters;
using robot.RiskTableAdapters;
using static robot.DataSet1;
using System.Threading.Tasks;
using robot.Parsers;

namespace robot
{
    class cl_Parser_MKD : cl_Parser
    {
        //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
        //SP sp = new SP();
        //SPRisk sprisk = new SPRisk();
        //DateTime reestr_date;
        //string report;
        //string pathFile;
        //int success = 0;

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
            int lastUsedRow = last.Row; // Последняя строка в документе
            MKD_DCA_rawDataTable mkd_dca = new MKD_DCA_rawDataTable();


            int i = 2; // Строка начала периода

            try
            {
                reestr_date = (DateTime)(sheet.Cells[i, 2] as Excel.Range).Value;
                reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);

                while (i <= lastUsedRow)
                {
                    MKD_DCA_rawRow mkd_dca_row = mkd_dca.NewMKD_DCA_rawRow();

                    mkd_dca_row["reestr_date"] = reestr_date;

                    mkd_dca_row["LN"] = (int)(sheet.Cells[i, 1] as Excel.Range).Value;
                    mkd_dca_row["Payment_date"] = (sheet.Cells[i, 2] as Excel.Range).Value;
                    mkd_dca_row["DCA_name"] = (sheet.Cells[i, 3] as Excel.Range).Value;
                    mkd_dca_row["Payment_amount"] = (double)(sheet.Cells[i, 4] as Excel.Range).Value;
                    mkd_dca_row["DCA_comission_amount"] = (double)(sheet.Cells[i, 5] as Excel.Range).Value;

                    mkd_dca.AddMKD_DCA_rawRow(mkd_dca_row);
                    mkd_dca.AcceptChanges();

                    Console.WriteLine((i - 1).ToString() + "/" + (lastUsedRow - 1).ToString() + " row uploaded");

                    i++;
                }

                if (mkd_dca.Rows.Count > 0)
                {
                    MKD_DCA_rawTableAdapter ad_MKD_DCA_raw = new MKD_DCA_rawTableAdapter();
                    ad_MKD_DCA_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"));

                    try
                    {
                        sp.sp_MKD_DCA_raw(mkd_dca);
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_DCA", "MKD", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_desc: " + exc.Message.ToString());
                        ex.Quit();
                    }
                }
                else
                {
                    report = "File was empty. There is no one row.";
                    logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_DCA", "MKD", DateTime.Now, false, report);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_desc: " + report);
                    ex.Quit();
                    return;
                }
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_DCA", "MKD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
                ex.Quit();
                return;
            }

            ex.Quit();


            report = "Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.";
            logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_DCA", "MKD", DateTime.Now, true, report);
            Console.WriteLine(report);

            TotalDcaForming();

            Console.WriteLine("Do you want to transport DCA to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();
            

            if (reply.Equals("Y"))
            {
                success = TransportDCAToRisk();
            }

            if (success == 1)
            {
                send_report = new cl_Send_Report("MKD_DCA", 1);
                //Console.WriteLine("Report was sended.");
            }

        }

        private void TotalDcaForming()
        {
            try
            {
                sp.sp_MKD_TOTAL_DCA(reestr_date);

                report = "[DWH_Risk].[dbo].[TOTAL_DCA] was formed.";
                logAdapter.InsertRow("cl_Parser_MKD", "TotalDcaForming", "MKD", DateTime.Now, false, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MKD", "TotalDcaForming", "MKD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }
        }

        private int TransportDCAToRisk()
        {
            try
            {
                sprisk.sp_MKD_TOTAL_DCA(reestr_date);
                report = "DCA was transported to [Risk].[dbo].[TOTAL_DCA]";
                logAdapter.InsertRow("cl_Parser_MKD", "TransportDCAToRisk", "MKD", DateTime.Now, true, report);
                Console.WriteLine(report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MKD", "TransportDCAToRisk", "MKD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }

        }

        public void parse_MKD_SNAP(Excel.Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_SNAP", "MKD", DateTime.Now, true, report);

            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastUsedRow = last.Row; // Последняя строка в документе

            new cl_Field_mapping(sheet, "loan", out int loan);
            new cl_Field_mapping(sheet, "current status", out int current_status);
            new cl_Field_mapping(sheet, "historical loan status", out int historical_loan_status);
            new cl_Field_mapping(sheet, "loan disbursment date", out int loan_disbursment_date);
            new cl_Field_mapping(sheet, "product", out int product);
            new cl_Field_mapping(sheet, "dpd", out int dpd);
            new cl_Field_mapping(sheet, "principal balance", out int principal_balance);
            new cl_Field_mapping(sheet, "monthly fee balance", out int monthly_fee_balance);
            new cl_Field_mapping(sheet, "guarantor fee balance", out int guarantor_fee_balance);
            new cl_Field_mapping(sheet, "penalty fee balance", out int penalty_fee_balance);
            new cl_Field_mapping(sheet, "penalty interest balance", out int penalty_interest_balance);
            new cl_Field_mapping(sheet, "interest balance", out int interest_balance);

            int i = 2; // Строка начала периода

            try
            {
                fileName = ex.Workbooks.Item[1].Name;
                fileName = fileName.Substring(fileName.IndexOf("_") + 1, 10).Replace("+", ""); //.ToString("yyyy-MM-dd");

                reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Excel.Range).Value;
                reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);

                MKD_SNAP_rawDataTable mkd_snap_raw = new MKD_SNAP_rawDataTable();
                System.Data.DataTable mkd_snap = new System.Data.DataTable();
                for (int j = 0; j < mkd_snap_raw.Columns.Count; j++)
                    mkd_snap.Columns.Add(mkd_snap_raw.Columns[j].ColumnName, mkd_snap_raw.Columns[j].DataType);

                while (i <= lastUsedRow)
                {
                    //MKD_SNAP_rawRow mkd_snap_row = mkd_snap.NewMKD_SNAP_rawRow();
                    System.Data.DataRow mkd_snap_row = mkd_snap.NewRow();

                    mkd_snap_row["reestr_date"] = reestr_date;

                    mkd_snap_row["Loan"] = (sheet.Cells[i, loan] as Excel.Range).Value.ToString();
                    mkd_snap_row["Current_status"] = (sheet.Cells[i, current_status] as Excel.Range).Value;
                    mkd_snap_row["Loan_disbursement_date"] = DateTime.Parse((sheet.Cells[i, loan_disbursment_date] as Excel.Range).Value);
                    mkd_snap_row["Product"] = (sheet.Cells[i, product] as Excel.Range).Value;
                    mkd_snap_row["DPD"] = (int)(sheet.Cells[i, dpd] as Excel.Range).Value;
                    mkd_snap_row["Historical_loan_status"] = (sheet.Cells[i, historical_loan_status] as Excel.Range).Value;
                    mkd_snap_row["Principal_balance"] = (double)(sheet.Cells[i, principal_balance] as Excel.Range).Value;
                    mkd_snap_row["Monthly_fee_balance"] = (double)(sheet.Cells[i, monthly_fee_balance] as Excel.Range).Value;
                    mkd_snap_row["Guarantor_fee_balance"] = (double)(sheet.Cells[i, guarantor_fee_balance] as Excel.Range).Value;
                    mkd_snap_row["Penalty_fee_balance"] = (double)(sheet.Cells[i, penalty_fee_balance] as Excel.Range).Value;
                    mkd_snap_row["Penalty_interest_balance"] = (double)(sheet.Cells[i, penalty_interest_balance] as Excel.Range).Value;
                    mkd_snap_row["Interest_balance"] = (double)(sheet.Cells[i, interest_balance] as Excel.Range).Value;

                    //mkd_snap.AddMKD_SNAP_rawRow(mkd_snap_row);
                    mkd_snap.Rows.Add(mkd_snap_row);
                    mkd_snap.AcceptChanges();

                    Console.WriteLine((i - 1).ToString() + "/" + (lastUsedRow - 1).ToString() + " row uploaded");

                    i++;
                }

                if (mkd_snap.Rows.Count > 0)
                {
                    MKD_SNAP_rawTableAdapter ad_MKD_SNAP_raw = new MKD_SNAP_rawTableAdapter();
                    ad_MKD_SNAP_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"));

                    try
                    {
                        task = new cl_Tasks("exec DWH_Risk.dbo.sp_MKD_SNAP_raw @MKD_SNAP_raw = ", mkd_snap);
                        //sp.sp_MKD_SNAP_raw(mkd_snap);
                    }
                    catch (Exception exc)
                    {
                        COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                        logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_SNAP", "MKD", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_desc: " + exc.Message.ToString());
                        ex.Quit();

                        return;
                    }

                    report = "Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.";
                    Console.WriteLine(report);
                    logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_SNAP", "MKD", DateTime.Now, true, report);

                    TotalSnapForming();
                }
                else
                {
                    report = "File was empty. There is no one row.";
                    logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_SNAP", "MKD", DateTime.Now, false, report);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_desc: " + report);
                    ex.Quit();

                    return;
                }
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MKD", "parse_MKD_SNAP", "MKD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
                ex.Quit();

                return;
            }


            ex.Quit();

            //report = "Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.";

            Console.WriteLine("Do you want to transport Snap to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                TransportSnapToRisk(reestr_date);
                TransportSnapCFToRisk(reestr_date);
            }

            send_report = new cl_Send_Report("MKD_SNAP", 1);
            //Console.WriteLine("Report was sended.");

        }

        private void TotalSnapForming()                     //to do: insert into try-catch
        {
            sp.sp_MKD2_portfolio_snapshot();

            try
            {
                task = new cl_Tasks("exec DWH_Risk.dbo.sp_MKD_TOTAL_SNAP");

                report = "[DWH_Risk].[dbo].[MKD2_portfolio_snapshot], [DWH_Risk].[dbo].[TOTAL_SNAP] were formed.";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_MKD", "TotalSnapForming", "MKD", DateTime.Now, true, report);
            }
            catch (Exception ex)
            {
                report = ex.Message;
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_MKD", "TotalSnapForming", "MKD", DateTime.Now, false, report);
            }

            try
            {
                task = new cl_Tasks("exec DWH_Risk.dbo.sp_MKD_TOTAL_SNAP_CFIELD");

                report = "[DWH_Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_MKD", "TotalSnapForming", "MKD", DateTime.Now, true, report);
            }
            catch (Exception ex)
            {
                report = ex.Message;
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_MKD", "TotalSnapForming", "MKD", DateTime.Now, false, report);
            }

            //sp.sp_MKD_TOTAL_SNAP_CFIELD();

        }

        private void TransportSnapToRisk(DateTime snapdate)
        {
            //Task task_snap = new Task(() =>
            //{
            //    sprisk.sp_MKD_TOTAL_SNAP(snapdate);
            //},
            //TaskCreationOptions.LongRunning);

            try
            {
                //task_snap.RunSynchronously();
                task = new cl_Tasks("exec Risk.dbo.sp_MKD_TOTAL_SNAP @date = '" + snapdate.ToString("yyyy-MM-dd") + "'");

                report = "Snap was transported to [Risk].[dbo].[MKD2_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP].";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_MKD", "TransportSnapToRisk", "MKD", DateTime.Now, true, report);

            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MKD", "TransportSnapToRisk", "MKD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }

        }

        private void TransportSnapCFToRisk(DateTime snapdate)
        {
            //Task task_snap_cf = new Task(() =>
            //{
            //    sprisk.sp_MKD_TOTAL_SNAP_CFIELD(snapdate);
            //},
            //TaskCreationOptions.LongRunning);

            try
            {
                //task_snap_cf.RunSynchronously();
                task = new cl_Tasks("exec Risk.dbo.sp_MKD_TOTAL_SNAP_CFIELD @date = '" + snapdate.ToString("yyyy-MM-dd") + "'");

                report = "[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_MKD", "TransportSnapToRisk", "MKD", DateTime.Now, true, report);

                //report into log
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MKD", "TransportSnapToRisk", "MKD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }

        }

    }
}