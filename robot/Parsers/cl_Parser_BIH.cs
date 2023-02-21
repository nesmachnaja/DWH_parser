using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using robot.DataSet1TableAdapters;
using Microsoft.Office.Interop.Excel;
using robot.RiskTableAdapters;
using robot.Total_BosniaTableAdapters;
using System.Threading.Tasks;
using static robot.DataSet1;
using robot.Parsers;

namespace robot
{
    class cl_Parser_BIH : cl_Parser
    {
        BIH_DCA_rawDataTable bih_dca_raw = new BIH_DCA_rawDataTable();
        System.Data.DataTable bih_dca = new System.Data.DataTable();
        //cl_Tasks task;

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

        private void OpenFile(string pathFile)
        {
            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
            Application ex = new Application();
            Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл

            if (pathFile.ToLower().Contains("external_collection")) parse_BIH_DCA(ex);
            if (pathFile.Contains("snapshot")) parse_BIH_SNAP(ex);
        }

        public void parse_BIH_DCA(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_DCA", "BIH", DateTime.Now, true, report);

            lastUsedRow = 0;
            string fileName = ex.Workbooks.Item[1].Name;
            
            //int startIndex = fileName.LastIndexOf("_") + 1;
            fileName = "01." + fileName.ToLower().Replace(".xlsx","").Replace("external_collection_","").Replace("_",".");
            reestr_date = DateTime.Parse(fileName).AddMonths(1).AddDays(-1);

            for (int j = 0; j < bih_dca_raw.Columns.Count; j++)
                bih_dca.Columns.Add(bih_dca_raw.Columns[j].ColumnName, bih_dca_raw.Columns[j].DataType);

            for (int j = 1; j <= 2; j++)
            {
                Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(j); // берем первый лист;
                Console.WriteLine("Sheet #" + j.ToString());
                parse_BIH_DCA_current_sheet(sheet);
            }

            report = "Loading is ready. " + lastUsedRow.ToString() + " rows were processed.";
            Console.WriteLine(report);
            logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_DCA", "BIH", DateTime.Now, true, report);

            try
            {
                task = new cl_Tasks("exec DWH_Risk.dbo.sp_BIH2_DCA @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                task = new cl_Tasks("exec DWH_Risk.dbo.sp_BIH_TOTAL_DCA @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //sp.sp_BIH2_DCA(reestr_date);
                //sp.sp_BIH_TOTAL_DCA(reestr_date);
                report = "[DWH_Risk].[dbo].[BIH2_DCA] and [DWH_Risk].[dbo].[TOTAL_DCA] were formed.";
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_DCA", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
                ex.Quit();

                return;
            }

            ex.Quit();

            Console.WriteLine("Do you want to transport DCA to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                success = TransportDCAToRisk();
            }

            if (success == 1)
            {
                //cl_PQR_Forming pqr = new cl_PQR_Forming("BIH");
                send_report = new cl_Send_Report("BIH_DCA", 1);
            }
        }

        private int TransportDCAToRisk()
        {
            /*Task task = new Task(() =>
            {
                SPRisk sprisk = new SPRisk();
                sprisk.sp_BIH_TOTAL_DCA(reestr_date);
            },
            TaskCreationOptions.LongRunning);*/

            try
            {
                task = new cl_Tasks("exec Risk.dbo.sp_BIH_TOTAL_DCA @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //task.RunSynchronously();

                Console.WriteLine("DCA was transported to [Risk].[dbo].[BIH2_DCA], [Risk].[dbo].[TOTAL_DCA]");
                report = "DCA was transported to [Risk].[dbo].[BIH2_DCA], [Risk].[dbo].[TOTAL_DCA]";
                logAdapter.InsertRow("cl_Parser_BIH", "TransportDCAToRisk", "BIH", DateTime.Now, true, report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_BIH", "TransportDCAToRisk", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }

        }

        private void parse_BIH_DCA_current_sheet(Worksheet sheet)
        {
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range range = sheet.get_Range("A1", last);
            int lastUsedColumn = last.Column;

            int firstNull = 0;

            for (int firstEmpty = 1; firstEmpty < last.Row; firstEmpty++)
            {
                if (sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]) == 0)
                {
                    firstNull = firstEmpty;
                    break;
                }
            }

            lastUsedRow = lastUsedRow + firstNull - 2; // Последняя строка в документе

            int i = 2; // Строка начала периода

            try
            {
                string debt_collector = (sheet.Cells[i, 5] as Range).Value;

                while (i < firstNull)
                {
                    System.Data.DataRow bih_dca_row = bih_dca.NewRow();
                    //BIH_DCA_rawRow bih_dca_row = bih_dca.NewBIH_DCA_rawRow();

                    bih_dca_row["Reestr_date"] = reestr_date;

                    bih_dca_row["Loan"] = (sheet.Cells[i, 1] as Range).Value.ToString();
                    bih_dca_row["Client"] = (sheet.Cells[i, 2] as Range).Value;
                    bih_dca_row["DPD"] = (int)(sheet.Cells[i, 3] as Range).Value;
                    bih_dca_row["Bucket"] = (sheet.Cells[i, 4] as Range).Value;
                    bih_dca_row["Debt_collector"] = debt_collector;
                    bih_dca_row["Amount"] = (double)(sheet.Cells[i, 6] as Range).Value;
                    bih_dca_row["Percent"] = (double)(sheet.Cells[i, 7] as Range).Value;
                    bih_dca_row["Fee_amount"] = (double)(sheet.Cells[i, 8] as Range).Value;

                    //bih_dca.AddBIH_DCA_rawRow(bih_dca_row);
                    bih_dca.Rows.Add(bih_dca_row);
                    bih_dca.AcceptChanges();
                    
                    Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");

                    i++;
                }

                if (bih_dca.Rows.Count > 0)
                {
                    BIH_DCA_rawTableAdapter ad_BIH_DCA_raw = new BIH_DCA_rawTableAdapter();
                    ad_BIH_DCA_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"), debt_collector);


                    try
                    {
                        task = new cl_Tasks("exec DWH_Risk.dbo.sp_BIH_DCA_raw @BIH_DCA_raw = ", bih_dca);
                        //sp.sp_BIH_DCA_raw(bih_dca);
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_DCA_current_sheet", "BIH", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_desc: " + exc.Message.ToString());
                        sheet.Application.Quit();

                        return;
                    }
                }
                else
                {
                    report = "File was empty. There is no one row.";
                    logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_DCA_current_sheet", "BIH", DateTime.Now, false, report);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_desc: " + report);
                    sheet.Application.Quit();

                    return;
                }

            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_DCA_current_sheet", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
                sheet.Application.Quit();

                return;
            }

        }


        public void parse_BIH_SNAP(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_SNAP", "BIH", DateTime.Now, true, report);
            
            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item("Report"); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            lastUsedRow = last.Row; // Последняя строка в документе
            BIH_SNAP_rawDataTable bih_snap_raw = new BIH_SNAP_rawDataTable();
            System.Data.DataTable bih_snap = new System.Data.DataTable();
            for (int j = 0; j < bih_snap_raw.Columns.Count; j++)
                bih_snap.Columns.Add(bih_snap_raw.Columns[j].ColumnName, bih_snap_raw.Columns[j].DataType);

            int i = 2; // Строка начала периода

            try
            {
                string fileName = ex.Workbooks.Item[1].Name;
                fileName = fileName.Substring(fileName.IndexOf("_") + 1, 10); //.ToString("yyyy-MM-dd");

                reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                //BIH_SNAP.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                //BIH_SNAP.Reestr_date = reestr_date;       //current date

                while (i <= lastUsedRow)
                {
                    //BIH_SNAP_rawRow bih_snap_raw = bih_snap.NewBIH_SNAP_rawRow();
                    System.Data.DataRow bih_snap_row = bih_snap.NewRow();

                    bih_snap_row["Reestr_date"] = reestr_date;

                    bih_snap_row["Loan"] = (sheet.Cells[i, 1] as Range).Value;
                    bih_snap_row["Client"] = (sheet.Cells[i, 2] as Range).Value;
                    bih_snap_row["Status"] = (sheet.Cells[i, 3] as Range).Value;
                    bih_snap_row["Loan_disbursment_date"] = DateTime.Parse((sheet.Cells[i, 4] as Range).Value);
                    bih_snap_row["Product"] = (sheet.Cells[i, 5] as Range).Value;
                    bih_snap_row["DPD"] = (int)(sheet.Cells[i, 6] as Range).Value;
                    bih_snap_row["Matured_principle"] = (double)(sheet.Cells[i, 7] as Range).Value;
                    bih_snap_row["Outstanding_principle"] = (double)(sheet.Cells[i, 8] as Range).Value;
                    bih_snap_row["Principal_balance"] = (double)(sheet.Cells[i, 9] as Range).Value;
                    bih_snap_row["Monthly_fee"] = (double)(sheet.Cells[i, 10] as Range).Value;
                    bih_snap_row["Guarantor_fee"] = (double)(sheet.Cells[i, 11] as Range).Value;
                    bih_snap_row["Penalty_fee"] = (double)(sheet.Cells[i, 12] as Range).Value;
                    bih_snap_row["Penalty_interest"] = (double)(sheet.Cells[i, 13] as Range).Value;
                    bih_snap_row["Interest_balance"] = (double)(sheet.Cells[i, 14] as Range).Value;
                    bih_snap_row["Credit_amount"] = (double)(sheet.Cells[i, 15] as Range).Value;
                    bih_snap_row["Available_limit"] = (double)(sheet.Cells[i, 16] as Range).Value;

                    //bih_snap.AddBIH_SNAP_rawRow(bih_snap_raw);
                    bih_snap.Rows.Add(bih_snap_row);
                    bih_snap.AcceptChanges();

                    Console.WriteLine((i - 1).ToString() + "/" + (lastUsedRow - 1).ToString() + " row uploaded");

                    i++;
                }

                if (bih_snap.Rows.Count > 0)
                {
                    BIH_SNAP_rawTableAdapter ad_BIH_SNAP_raw = new BIH_SNAP_rawTableAdapter();
                    ad_BIH_SNAP_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"));

                    /*Task task_snap = new Task(() =>
                    {
                        sp.sp_BIH_SNAP_raw(bih_snap);
                    },
                    TaskCreationOptions.LongRunning);*/

                    try
                    {
                        task = new cl_Tasks("exec DWH_Risk.dbo.sp_BIH_SNAP_raw @BIH_SNAP_raw = ", bih_snap);
                        //task_snap.RunSynchronously();
                        //sp.sp_BIH_SNAP_raw(bih_snap);
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_SNAP", "BIH", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_desc: " + exc.Message.ToString());
                        ex.Quit();

                        return;
                    }


                    report = "Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.";
                    Console.WriteLine(report);
                    logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_SNAP", "BIH", DateTime.Now, true, report);
                    
                    TotalSnapForming();

                    TotalSnapCFieldForming();
                    //report = "[DWH_Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                    //Console.WriteLine(report);
                    //logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_SNAP", "BIH", DateTime.Now, true, report);
                }
                else
                {
                    report = "File was empty. There is no one row.";
                    logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_SNAP", "BIH", DateTime.Now, false, report);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_desc: " + report);

                    return;
                }

            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_SNAP", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
                ex.Quit();

                return;
            }


            ex.Quit();

            TransportSnapToBosnia(reestr_date);

            Console.WriteLine("Do you want to transport Snap to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                TransportSnapToRisk();
                success = TransportSnapCFToRisk();
            }

            if (success == 1)
            {
                send_report = new cl_Send_Report("BIH_SNAP", 1);
                //Console.WriteLine("Report was sended.");
            }
        }

        private void TotalSnapForming()
        {
            try
            {
                task = new cl_Tasks("exec DWH_Risk.dbo.sp_BIH2_portfolio_snapshot @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //sp.sp_BIH2_portfolio_snapshot(reestr_date);
                report = "[DWH_Risk].[dbo].[BIH2_portfolio_snapshot] was formed.";
                logAdapter.InsertRow("cl_Parser_BIH", "TotalSnapForming", "BIH", DateTime.Now, true, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_BIH", "TotalSnapForming", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }

            try
            {
                task = new cl_Tasks("exec DWH_Risk.dbo.sp_BIH_TOTAL_SNAP @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //sp.sp_BIH_TOTAL_SNAP(reestr_date);
                report = "[DWH_Risk].[dbo].[TOTAL_SNAP] was formed.";
                logAdapter.InsertRow("cl_Parser_BIH", "TotalSnapForming", "BIH", DateTime.Now, true, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_BIH", "TotalSnapForming", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }
        }

        private void TransportSnapToBosnia(DateTime snapdate)
        {
            try
            {
                TOTAL_SNAPTableAdapter ad_TOTAL_SNAP = new TOTAL_SNAPTableAdapter();
                ad_TOTAL_SNAP.DeletePeriod(snapdate.ToString("yyyy-MM-dd"));
                ad_TOTAL_SNAP.InsertPeriod(snapdate);

                Console.WriteLine("Snap was transported to [Total_Bosnia].[dbo].[TOTAL_SNAP]");
                report = "Snap was transported to [Total_Bosnia].[dbo].[TOTAL_SNAP]";
                logAdapter.InsertRow("cl_Parser_BIH", "TransportSnapToBosnia", "BIH", DateTime.Now, true, report);
                //report into log

                //SPBosnia spbosnia = new SPBosnia();
                //SP sp = new SP();
                //sp.sp_BIH_TOTAL_SNAP_CFIELD();
                TOTAL_SNAP_CFIELDTableAdapter ad_TOTAL_SNAP_CFIELD = new TOTAL_SNAP_CFIELDTableAdapter();
                ad_TOTAL_SNAP_CFIELD.DeletePeriod(snapdate.ToString("yyyy-MM-dd"));
                ad_TOTAL_SNAP_CFIELD.InsertPeriod(snapdate);

                Console.WriteLine("CField was transported to [Total_Bosnia].[dbo].[TOTAL_SNAP_CFIELD].");
                report = "CField was transported to [Total_Bosnia].[dbo].[TOTAL_SNAP_CFIELD].";
                logAdapter.InsertRow("cl_Parser_BIH", "TransportSnapToBosnia", "BIH", DateTime.Now, true, report);

            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_BIH", "TransportSnapToBosnia", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }

            //Console.ReadKey();
        }

        private void TotalSnapCFieldForming()
        {
            try
            {
                task = new cl_Tasks("exec DWH_Risk.dbo.sp_BIH_TOTAL_SNAP_CFIELD");

                report = "[DWH_Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_BIH", "TotalSnapCFieldForming", "BIH", DateTime.Now, true, report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_BIH", "TotalSnapCFieldForming", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }

        }
        
        private void TransportSnapToRisk()
        {
            /*
            Task task_snap = new Task(() =>
            {
                SPRisk sprisk = new SPRisk();
                sprisk.sp_BIH_TOTAL_SNAP(reestr_date);
            },
            TaskCreationOptions.LongRunning);*/

            try
            {
                task = new cl_Tasks("exec Risk.dbo.sp_BIH_TOTAL_SNAP @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //task_snap.RunSynchronously();

                //SPRisk sprisk = new sp();
                //sprisk.sp_BIH_TOTAL_SNAP(snapdate);
                Console.WriteLine("Snap was transported to [Risk].[dbo].[BIH2_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP]");
                report = "Snap was transported to [Risk].[dbo].[BIH2_portfolio_snapshot], [Risk].[dbo].[TOTAL_SNAP]";
                logAdapter.InsertRow("cl_Parser_BIH", "TransportSnapToRisk", "BIH", DateTime.Now, true, report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_BIH", "TransportSnapToRisk", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }

            //Console.ReadKey();
        }

        private int TransportSnapCFToRisk()
        {
            //Task task_snapCF = new Task(() =>
            //{
            //    SPRisk sprisk = new SPRisk();
            //    sprisk.sp_BIH_TOTAL_SNAP_CFIELD(reestr_date);
            //},
            //TaskCreationOptions.LongRunning);

            try
            {
                task = new cl_Tasks("exec Risk.dbo.sp_BIH_TOTAL_SNAP_CFIELD @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //task_snapCF.RunSynchronously();

                //sprisk.sp_BIH_TOTAL_SNAP_CFIELD(snapdate);
                Console.WriteLine("Snap_CField was transported to [Risk].[dbo].[TOTAL_SNAP_CFIELD]");
                report = "Snap_CField was transported to [Risk].[dbo].[TOTAL_SNAP_CFIELD]";
                logAdapter.InsertRow("cl_Parser_BIH", "TransportSnapToRisk", "BIH", DateTime.Now, true, report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_BIH", "TransportSnapToRisk", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }

        }
    }
}