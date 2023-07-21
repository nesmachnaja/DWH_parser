using robot.DataSet1TableAdapters;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using robot.RiskTableAdapters;
using static robot.DataSet1;
using System.Text.RegularExpressions;

namespace robot.Parsers
{
    class cl_Parser_MD_test : cl_Parser
    {
        public void StartParsing(string path_file)
        {
            logAdapter = new COUNTRY_LogTableAdapter();
            int correctPath = 0;

            while (correctPath == 0)
            {
                try
                {
                    pathFile = path_file;
                    OpenFile();
                    correctPath = 1;
                }
                catch (Exception exc)
                {
                    Console.WriteLine(exc.Message);
                    Console.WriteLine("Incorrect file path.");
                }
            }
        }

        /*
        private static string GetPath()
        {
            Console.WriteLine("Appoint file path: ");
            string pathFile = Console.ReadLine();
            return pathFile;
        }*/

        public void OpenFile()
        {
            string fullPath = Path.GetFullPath(pathFile);
            Application ex = new Application();
            ex.DisplayAlerts = false;
            Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, 1);

            if (pathFile.Contains("Plati")) parse_MD_DCA(ex);
            if (pathFile.Contains("SNAP") || pathFile.Contains("WO")) parse_MD_SNAP(ex);
        }

        public void DcaPostProcessing()
        {
            TotalDCAForming();
                        
            success = TransportDCAToRisk();
            
            if (success == 1)
            {
                send_report = new cl_Send_Report("MD_DCA", 1);
                //Console.WriteLine("Report was sended.");
            }
        }

        public void SnapPostProcessing()
        {
            TransportMDSnapToRisk();
            TransportTotalSnapToRisk();
            success = TransportSnapCFToRisk();

            if (success == 1)
            {
                send_report = new cl_Send_Report("MD_SNAP", 1);
                //Console.WriteLine("Report was sended.");
            }
        }

        public void parse_MD_DCA(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA", "MD", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(2); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastUsedRow = last.Row; // Последняя строка в документе
            MD_DCA_rawDataTable md_dca_raw = new MD_DCA_rawDataTable();
            System.Data.DataTable md_dca = new System.Data.DataTable();
            for (int j = 0; j < md_dca_raw.Columns.Count; j++)
                md_dca.Columns.Add(md_dca_raw.Columns[j].ColumnName, md_dca_raw.Columns[j].DataType);

            int i = lastUsedRow; // Строка начала периода

            new cl_Field_mapping(sheet, "Collection company", out int сollection_company);
            new cl_Field_mapping(sheet, "Payment month", out int payment_month);
            new cl_Field_mapping(sheet, "Debtor", out int debtor);
            new cl_Field_mapping(sheet, "IDNP debitorului", out int idnp_debitorului);
            new cl_Field_mapping(sheet, "Contract no.", out int contract);
            new cl_Field_mapping(sheet, "Total paid", out int total_paid);
            new cl_Field_mapping(sheet, "Fee", out int fee);
            new cl_Field_mapping(sheet, "Fee (incl. VAT)", out int fee_including_vat);
            new cl_Field_mapping(sheet, "Payment date", out int payment_date);

            try
            {
                reestr_date = (DateTime)(sheet.Cells[i, 2] as Range).Value;
                //MD_DCA.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);

                while (i > 0)
                {
                    System.Data.DataRow md_dca_row = md_dca.NewRow();
                    //MD_DCA_rawRow md_dca_row = md_dca.NewMD_DCA_rawRow();

                    md_dca_row["reestr_date"] = reestr_date;

                    md_dca_row["Collection_company"] = (sheet.Cells[i, сollection_company] as Range).Value;
                    md_dca_row["Payment_month"] = (DateTime)(sheet.Cells[i, payment_month] as Range).Value;
                    md_dca_row["Debtor"] = (sheet.Cells[i, debtor] as Range).Value;
                    md_dca_row["IDNP_debitorului"] = (sheet.Cells[i, idnp_debitorului] as Range).Value;
                    md_dca_row["Contract"] = (sheet.Cells[i, contract] as Range).Value;
                    md_dca_row["Total_paid"] = (double)(sheet.Cells[i, total_paid] as Range).Value;
                    md_dca_row["Fee"] = (double)(sheet.Cells[i, fee] as Range).Value;
                    md_dca_row["Fee_including_VAT"] = (double)(sheet.Cells[i, fee_including_vat] as Range).Value;
                    //md_dca_row["Types"] = (sheet.Cells[i, 9] as Range).Value;
                    md_dca_row["Payment_date"] = DateTime.Parse((sheet.Cells[i, payment_date] as Range).Value.ToString().Replace("0:00:00", ""));

                    //if ((DateTime)md_dca_row["Payment_month"] != reestr_date)

                    //md_dca.AddMD_DCA_rawRow(md_dca_row);
                    md_dca.Rows.Add(md_dca_row);
                    md_dca.AcceptChanges();


                    if ((DateTime)md_dca_row["Payment_month"] != reestr_date)
                    {
                        Console.WriteLine("The other rows are marked by another Payment_month");

                        break;
                    }
                    else
                    {
                        Console.WriteLine((lastUsedRow - i + 1).ToString() + "/" + (lastUsedRow - 1).ToString() + " row uploaded");

                        //md_dca.AddMD_DCA_rawRow(md_dca_row);
                        //md_dca.AcceptChanges();
                    }

                    i--;

                }

                if (md_dca.Rows.Count > 0)
                {
                    MD_DCA_rawTableAdapter ad_MD_DCA_raw = new MD_DCA_rawTableAdapter();
                    ad_MD_DCA_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"));

                    try
                    {
                        task = new cl_Tasks("exec DWH_Risk.dbo.sp_MD_DCA_raw @MD_DCA_raw = ", md_dca);
                        //sp.sp_MD_DCA_raw(md_dca);

                        report = "Loading is ready. " + (lastUsedRow - i).ToString() + " rows were processed.";
                        logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA", "MD", DateTime.Now, true, report);
                        Console.WriteLine(report);
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA", "MD", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_desc: " + exc.Message.ToString());

                        return;
                    }
                }
                else
                {
                    report = "File was empty. There is no one row.";
                    logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA", "MD", DateTime.Now, false, report);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_descr: " + report);
                    ex.Quit();

                    return;
                }
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                ex.Quit();

                return;
            }

            ex.Quit();

        }

        private void TotalDCAForming()
        {
            try
            {
                sp.sp_MD2_DCA(reestr_date);

                report = "[dbo].[MD2_DCA] was formed.";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_MD", "TotalDCAForming", "MD", DateTime.Now, true, report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "TotalDCAForming", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }

            try
            {
                sp.sp_MD_TOTAL_DCA();

                report = "[dbo].[TOTAL_DCA] was formed.";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_MD", "TotalDCAForming", "MD", DateTime.Now, true, report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "TotalDCAForming", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }
        }

        public void parse_MD_SNAP(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_MD", "parse_MD_SNAP", "MD", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            lastUsedRow = last.Row; // Последняя строка в документе
            MD_SNAP_rawDataTable md_snap_raw = new MD_SNAP_rawDataTable();
            System.Data.DataTable md_snap = new System.Data.DataTable();
            for (int j = 0; j < md_snap_raw.Columns.Count; j++)
                md_snap.Columns.Add(md_snap_raw.Columns[j].ColumnName, md_snap_raw.Columns[j].DataType);


            int firstNull = SearchFirstNullRow(sheet, lastUsedRow);

            int i = 0;
            for (int start = 3; start < firstNull; start++)
                if (int.TryParse((sheet.Cells[start, 1] as Range).Text, out i))
                {
                    i = start;
                    break;
                }

            //int i = 3; // Строка начала периода
            int startPosition = i - 1; // Строка начала периода

            new cl_Field_mapping(sheet, "Account ID", out int account_id);
            new cl_Field_mapping(sheet, "Loan Amount", out int loan_amount);
            new cl_Field_mapping(sheet, "Days In Arrears", out int dpd);
            new cl_Field_mapping(sheet, "Principal Balance", out int principal_balance);
            new cl_Field_mapping(sheet, "Principal", out int principal);
            new cl_Field_mapping(sheet, "origin.Fee", out int origination_fee);
            new cl_Field_mapping(sheet, "origin.Fee IL", out int origination_fee_il);
            new cl_Field_mapping(sheet, "Interest balance for provisions", out int interest_balance_for_provisions);
            
            try
            {
                fileName = ex.Workbooks.Item[1].Name;
                fileName = fileName.Replace("Moldova_SNAP ", "").Replace("Moldova_WO ", "").Replace("Moldova_WO_accumulated_", "").Replace(".xlsx", "").Replace("_", " "); //.ToString("yyyy-MM-dd");

                string pattern = @"\d+\.\d+\.\d+";
                Match result = Regex.Match(fileName, pattern);
                fileName = result.ToString();

                //ex.Quit();

                reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                //MD_SNAP.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth

                string source_type = ex.Workbooks.Item[1].Name.Replace(".xlsx", "");

                while (i < firstNull)
                {
                    //MD_SNAP_rawRow md_snap_row = md_snap.NewMD_SNAP_rawRow();
                    System.Data.DataRow md_snap_row = md_snap.NewRow();

                    md_snap_row["reestr_date"] = reestr_date.ToString("yyyy-MM-dd");
                    md_snap_row["SnapDate"] = reestr_date.ToString("yyyy-MM-dd");

                    md_snap_row["Account_ID"] = (sheet.Cells[i, account_id] as Range).Value.ToString();
                    md_snap_row["Loan_amount"] = (double)(sheet.Cells[i, loan_amount] as Range).Value;
                    md_snap_row["DPD"] = (int)(sheet.Cells[i, dpd] as Range).Value;
                    md_snap_row["Principal_balance"] = (double)(sheet.Cells[i, principal_balance] as Range).Value;
                    md_snap_row["Principal"] = (double)(sheet.Cells[i, principal] as Range).Value;
                    md_snap_row["Origination_fee"] = (double)(sheet.Cells[i, origination_fee] as Range).Value;
                    md_snap_row["Origination_fee_IL"] = (double)(sheet.Cells[i, origination_fee_il] as Range).Value;
                    md_snap_row["Interest_balance_for_provisions"] = (double)(sheet.Cells[i, interest_balance_for_provisions] as Range).Value;

                    md_snap_row["source_type"] = source_type;

                    //md_snap.AddMD_SNAP_rawRow(md_snap_row);
                    md_snap.Rows.Add(md_snap_row);
                    md_snap.AcceptChanges();

                    Console.WriteLine((i - startPosition).ToString() + "/" + (firstNull - startPosition - 1).ToString() + " row uploaded");

                    i++;
                }

                if (md_snap.Rows.Count > 0)
                {
                    MD_SNAP_rawTableAdapter ad_MD_SNAP_raw = new MD_SNAP_rawTableAdapter();
                    ad_MD_SNAP_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"), source_type);

                    try
                    {
                        task = new cl_Tasks("exec DWH_Risk.dbo.sp_MD_SNAP_raw @MD_SNAP_raw = ", md_snap);
                        //sp.sp_MD_SNAP_raw(md_snap);
                        //ad_MD_SNAP_raw.UpdateInitialsAndClients();

                        report = "Loading is ready. " + (firstNull - startPosition - 1).ToString() + " rows were processed.";
                        logAdapter.InsertRow("cl_Parser_MD", "parse_MD_SNAP", "MD", DateTime.Now, true, report);
                        Console.WriteLine(report);
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_MD", "parse_MD_SNAP", "MD", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_descr: " + exc.Message.ToString());
                        ex.Quit();

                        return;
                    }
                }
                else
                {
                    report = "File was empty. There is no one row.";
                    logAdapter.InsertRow("cl_Parser_MD", "parse_MD_SNAP", "MD", DateTime.Now, false, report);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_descr: " + report);
                    ex.Quit();

                    return;
                }
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "parse_MD_SNAP", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message.ToString());
                ex.Quit();

                return;
            }

            ex.ActiveWorkbook.RefreshAll();
            try
            {
                //ex.ActiveWorkbook.SaveAs();
                //ex.DisplayAlerts = false;
                ex.ActiveWorkbook.SaveAs(pathFile, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                ex.ActiveWorkbook.Close();
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                throw e;
            }

            ex.Quit();

        }

        private void TransportMDSnapToRisk()
        {
            /*Task task_md2_sn = new Task(() =>
            {
                sprisk.sp_MD2_portfolio_snapshot(reestr_date);
            },
            TaskCreationOptions.LongRunning);*/

            try
            {
                task = new cl_Tasks("exec Risk.dbo.sp_MD2_portfolio_snapshot @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //task_md2_sn.RunSynchronously();

                report = "Snap was transported to [Risk].[dbo].[MD2_portfolio_snapshot]";
                logAdapter.InsertRow("cl_Parser_MD", "TransportMDSnapToRisk", "MD", DateTime.Now, true, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "TransportMDSnapToRisk", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }

            //Task task_md3_sn = new Task(() =>
            //{
            //    sprisk.sp_MD3_portfolio_snapshot(reestr_date);
            //},
            //TaskCreationOptions.LongRunning);

            try
            {
                //task_md3_sn.RunSynchronously();
                task = new cl_Tasks("exec Risk.dbo.sp_MD3_portfolio_snapshot @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");

                report = "IL-block was calculated in [Risk].[dbo].[MD3_portfolio_snapshot]";
                logAdapter.InsertRow("cl_Parser_MD", "TransportMDSnapToRisk", "MD", DateTime.Now, true, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "TransportMDSnapToRisk", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }

        }

        private void TransportTotalSnapToRisk()
        {

            /*Task task_total_snap = new Task(() =>
            {
                sprisk.sp_MD_TOTAL_SNAP(reestr_date);
            },
            TaskCreationOptions.LongRunning);*/

            try
            {
                task = new cl_Tasks("exec Risk.dbo.sp_MD_TOTAL_SNAP @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //task_total_snap.RunSynchronously();

                report = "[Risk].[dbo].[TOTAL_SNAP] was formed.";
                logAdapter.InsertRow("cl_Parser_MD", "TransportTotalSnapToRisk", "MD", DateTime.Now, true, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "TransportTotalSnapToRisk", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return;
            }

        }

        private int TransportSnapCFToRisk()
        {
            //Task task_snap_cf = new Task(() =>
            //{
            //    sprisk.sp_MD_TOTAL_SNAP_CFIELD();
            //},
            //TaskCreationOptions.LongRunning);

            try
            {
                //task_snap_cf.RunSynchronously();
                task = new cl_Tasks("exec Risk.dbo.sp_MD_TOTAL_SNAP_CFIELD");

                report = "[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                logAdapter.InsertRow("cl_Parser_MD", "TransportSnapCFToRisk", "MD", DateTime.Now, true, report);
                Console.WriteLine(report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "TransportSnapCFToRisk", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }

        }

        private int TransportDCAToRisk()
        {
            /*Task task_md_dca = new Task(() =>
            {
                sprisk.sp_MD_TOTAL_DCA(reestr_date);
            },
            TaskCreationOptions.LongRunning);*/

            try
            {
                task = new cl_Tasks("exec Risk.dbo.sp_MD_TOTAL_DCA @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //task_md_dca.RunSynchronously();

                report = "DCA was transported to [Risk].[dbo].[MD2_DCA], [Risk].[dbo].[TOTAL_DCA]";
                logAdapter.InsertRow("cl_Parser_MD", "TransportDCAToRisk", "MD", DateTime.Now, true, report);
                Console.WriteLine(report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "TransportDCAToRisk", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }

        }

        private static int SearchFirstNullRow(Worksheet sheet, int lastUsedRow)
        {
            if (sheet.Application.WorksheetFunction.CountA(sheet.Rows[lastUsedRow]) != 0
                && sheet.Application.WorksheetFunction.CountA(sheet.Rows[lastUsedRow]) >= sheet.Application.WorksheetFunction.CountA(sheet.Rows[5]) * 0.9)
                return lastUsedRow;

            int midpoint = lastUsedRow / 2;
            int firstNull = 0;

            //int n = (int)sheet.Application.WorksheetFunction.CountA(sheet.Rows[midpoint]);
            //int u = (int)sheet.Application.WorksheetFunction.CountA(sheet.Rows[5]);

            if (sheet.Application.WorksheetFunction.CountA(sheet.Rows[midpoint]) != 0
                && sheet.Application.WorksheetFunction.CountA(sheet.Rows[midpoint]) >= sheet.Application.WorksheetFunction.CountA(sheet.Rows[5]) * 0.9)
            {
                for (int firstEmpty = midpoint; firstEmpty <= lastUsedRow + 1; firstEmpty++)
                {
                    if (sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]) == 0
                    || sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]) < sheet.Application.WorksheetFunction.CountA(sheet.Rows[5]) * 0.9)
                    {
                        firstNull = firstEmpty;
                        break;
                    }
                }
            }
            else
            {
                for (int firstEmpty = midpoint; firstEmpty > 0; firstEmpty--)
                {
                    //int a = (int)sheet.Application.WorksheetFunction.CountBlank(sheet.Rows[firstEmpty]);
                    //int s = (int)sheet.Application.WorksheetFunction.CountBlank(sheet.Rows[5]);
                    if (sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]) != 0
                    && sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]) >= sheet.Application.WorksheetFunction.CountA(sheet.Rows[5]) * 0.9)
                    {
                        firstNull = firstEmpty + 1;
                        break;
                    }
                }
            }

            return firstNull;
        }

    }
}
