﻿using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using robot.DataSet1TableAdapters;
using Microsoft.Office.Interop.Excel;

namespace robot
{
    class cl_Parser_BIH
    {
        private int lastUsedRow;
        cl_BIH_DCA BIH_DCA = new cl_BIH_DCA();
        COUNTRY_LogTableAdapter logAdapter;

        public void OpenFile()
        {
            logAdapter = new COUNTRY_LogTableAdapter();

            string pathFile = @"C:\Users\Людмила\source\repos\robot\external_collection_03_2022.xlsx"; // Путь к файлу отчета
            //static string pathFile = @"C:\Users\Людмила\source\repos\robot\DCA.xlsx"; // Путь к файлу отчета
            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
            Application ex = new Application();
            Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл

            if (pathFile.Contains("external_collection")) parse_BIH_DCA(ex);
            if (pathFile.Contains("snapshot")) parse_BIH_SNAP(ex);
        }

        public void parse_BIH_DCA(Application ex)
        {
            string report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_DCA", "BIH", DateTime.Now, true, report);

            lastUsedRow = 0;
            string fileName = ex.Workbooks.Item[1].Name;
            
            //int startIndex = fileName.LastIndexOf("_") + 1;
            fileName = "01." + fileName.Replace(".xlsx","").Replace("external_collection_","").Replace("_",".");
            BIH_DCA.Reestr_date = DateTime.Parse(fileName).AddMonths(1).AddDays(-1);

            for (int j = 1; j <= 2; j++)
            {
                Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(j); // берем первый лист;
                Console.WriteLine("Sheet #" + j.ToString());
                parse_BIH_DCA_current_sheet(sheet);
            }

            try
            {
                SP sp = new SP();
                sp.sp_BIH2_DCA(BIH_DCA.Reestr_date);
                sp.sp_BIH_TOTAL_DCA(BIH_DCA.Reestr_date);
                Console.WriteLine("Loading is ready. " + lastUsedRow.ToString() + " rows were processed.");
            }
            catch (Exception exc)
            {
                //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_DCA", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                ex.Quit();
            }

            ex.Quit();

            report = "Loading is ready. " + lastUsedRow.ToString() + " rows were processed.";
            logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_DCA", "BIH", DateTime.Now, true, report);

            Console.ReadKey();

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
                BIH_DCA.Debt_collector = (sheet.Cells[i, 5] as Range).Value;

                BIH_DCA_rawTableAdapter ad_BIH_DCA_raw = new BIH_DCA_rawTableAdapter();
                ad_BIH_DCA_raw.DeletePeriod(BIH_DCA.Reestr_date.ToString("yyyy-MM-dd"), BIH_DCA.Debt_collector);

                while (i < firstNull)
                {
                    BIH_DCA.Loan = (sheet.Cells[i, 1] as Range).Value.ToString();
                    BIH_DCA.Client = (sheet.Cells[i, 2] as Range).Value;
                    BIH_DCA.DPD = (int)(sheet.Cells[i, 3] as Range).Value;
                    BIH_DCA.Bucket = (sheet.Cells[i, 4] as Range).Value;
                    BIH_DCA.Amount = (double)(sheet.Cells[i, 6] as Range).Value;
                    BIH_DCA.Percent = (double)(sheet.Cells[i, 7] as Range).Value;
                    BIH_DCA.Fee_amount = (double)(sheet.Cells[i, 8] as Range).Value;

                    try
                    {
                        ad_BIH_DCA_raw.InsertRow(BIH_DCA.Reestr_date.ToString("yyyy-MM-dd"), BIH_DCA.Loan, BIH_DCA.Client, BIH_DCA.DPD, BIH_DCA.Bucket, BIH_DCA.Debt_collector, BIH_DCA.Amount, BIH_DCA.Percent, BIH_DCA.Fee_amount);
                        Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");
                    }
                    catch (Exception exc)
                    {
                        //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                        logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_DCA_current_sheet", "BIH", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        sheet.Application.Quit();
                    }

                    i++;
                }

                
            }
            catch (Exception exc)
            {
                //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_DCA_current_sheet", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                sheet.Application.Quit();
            }



            //report                                                           ----TO_DO
        }


        public void parse_BIH_SNAP(Application ex)
        {
            string report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_SNAP", "BIH", DateTime.Now, true, report);
            
            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range range = sheet.get_Range("A1", last);
            lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;

            int i = 2; // Строка начала периода

            try
            {

                cl_BIH_SNAP BIH_SNAP = new cl_BIH_SNAP();

                string fileName = ex.Workbooks.Item[1].Name;
                fileName = fileName.Substring(fileName.IndexOf("_") + 1, 10); //.ToString("yyyy-MM-dd");

                DateTime reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                //BIH_SNAP.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                BIH_SNAP.Reestr_date = reestr_date;       //current date

                BIH_SNAP_rawTableAdapter ad_BIH_SNAP_raw = new BIH_SNAP_rawTableAdapter();
                ad_BIH_SNAP_raw.DeletePeriod(BIH_SNAP.Reestr_date.ToString("yyyy-MM-dd"));

                while (i <= lastUsedRow)
                {
                    BIH_SNAP.Loan = (sheet.Cells[i, 1] as Range).Value;
                    BIH_SNAP.Client = (sheet.Cells[i, 2] as Range).Value;
                    BIH_SNAP.Status = (sheet.Cells[i, 3] as Range).Value;
                    BIH_SNAP.Loan_disbursment_date = DateTime.Parse((sheet.Cells[i, 4] as Range).Value);
                    BIH_SNAP.Product = (sheet.Cells[i, 5] as Range).Value;
                    BIH_SNAP.DPD = (int)(sheet.Cells[i, 6] as Range).Value;
                    BIH_SNAP.Matured_principle = (double)(sheet.Cells[i, 7] as Range).Value;
                    BIH_SNAP.Outstanding_principle = (double)(sheet.Cells[i, 8] as Range).Value;
                    BIH_SNAP.Principal_balance = (double)(sheet.Cells[i, 9] as Range).Value;
                    BIH_SNAP.Monthly_fee = (double)(sheet.Cells[i, 10] as Range).Value;
                    BIH_SNAP.Guarantor_fee = (double)(sheet.Cells[i, 11] as Range).Value;
                    BIH_SNAP.Penalty_fee = (double)(sheet.Cells[i, 12] as Range).Value;
                    BIH_SNAP.Penalty_interest = (double)(sheet.Cells[i, 13] as Range).Value;
                    BIH_SNAP.Interest_balance = (double)(sheet.Cells[i, 14] as Range).Value;
                    BIH_SNAP.Credit_amount = (double)(sheet.Cells[i, 15] as Range).Value;
                    BIH_SNAP.Available_limit = (double)(sheet.Cells[i, 16] as Range).Value;

                    try
                    {
                        ad_BIH_SNAP_raw.InsertRow(BIH_SNAP.Reestr_date.ToString("yyyy-MM-dd"), BIH_SNAP.Loan, BIH_SNAP.Client, BIH_SNAP.Status, BIH_SNAP.Loan_disbursment_date.ToString("yyyy-MM-dd"),
                            BIH_SNAP.Product, BIH_SNAP.DPD, BIH_SNAP.Matured_principle, BIH_SNAP.Outstanding_principle, BIH_SNAP.Principal_balance, BIH_SNAP.Monthly_fee,
                            BIH_SNAP.Guarantor_fee, BIH_SNAP.Penalty_fee, BIH_SNAP.Penalty_interest, BIH_SNAP.Interest_balance, BIH_SNAP.Credit_amount, BIH_SNAP.Available_limit);
                        Console.WriteLine((i - 1).ToString() + "/" + (lastUsedRow - 1).ToString() + " row uploaded");
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_SNAP", "BIH", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        ex.Quit();
                    }

                    i++;
                }

                SP sp = new SP();
                sp.sp_BIH2_portfolio_snapshot(BIH_SNAP.Reestr_date);
                sp.sp_BIH_TOTAL_SNAP(BIH_SNAP.Reestr_date);

                Console.WriteLine("Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.");
            }
            catch (Exception exc)
            {
                //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_SNAP", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                ex.Quit();
            }


            ex.Quit();

            report = "Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.";
            logAdapter.InsertRow("cl_Parser_BIH", "parse_BIH_SNAP", "BIH", DateTime.Now, true, report);

            Console.ReadKey();

            //report                                                           ----TO_DO

        }


    }
}