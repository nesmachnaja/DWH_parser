using System;
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

        public void OpenFile()
        {
            string pathFile = @"C:\Users\Людмила\source\repos\robot\external_collection_02.2022.xlsx"; // Путь к файлу отчета
            //static string pathFile = @"C:\Users\Людмила\source\repos\robot\DCA.xlsx"; // Путь к файлу отчета
            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
            Excel.Application ex = new Excel.Application();
            Excel.Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл

            if (pathFile.Contains("external_collection")) parse_BIH_DCA(ex);
            //if (pathFile.Contains("SNAP_")) parse_BIH_SNAP(ex);
        }

        public void parse_BIH_DCA(Excel.Application ex)
        {
            lastUsedRow = 0;
            string fileName = ex.Workbooks.Item[1].Name;
            
            int startIndex = fileName.LastIndexOf("_") + 1;
            fileName = "01." + fileName.Substring(startIndex, fileName.Length - startIndex).Replace(".xlsx","");
            BIH_DCA.Reestr_date = DateTime.Parse(fileName).AddMonths(1).AddDays(-1);

            for (int j = 1; j <= 2; j++)
            {
                Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(j); // берем первый лист;
                parse_BIH_DCA_current_sheet(sheet);
            }

            try
            {
                SP sp = new SP();
                sp.sp_BIH2_DCA(BIH_DCA.Reestr_date);
                sp.sp_BIH_TOTAL_DCA(BIH_DCA.Reestr_date);
                Console.WriteLine("Loading is ready. " + (lastUsedRow).ToString() + " rows were processed.");
            }
            catch (Exception exc)
            {
                COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_BIH", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
            }
            
            ex.Quit();

            Console.ReadKey();

        }

        private void parse_BIH_DCA_current_sheet(Excel.Worksheet sheet)
        {
            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = sheet.get_Range("A1", last);
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
                BIH_DCA.Debt_collector = (sheet.Cells[i, 5] as Excel.Range).Value;

                BIH_DCA_rawTableAdapter ad_BIH_DCA_raw = new BIH_DCA_rawTableAdapter();
                ad_BIH_DCA_raw.DeletePeriod(BIH_DCA.Reestr_date.ToString("yyyy-MM-dd"), BIH_DCA.Debt_collector);

                while (i < firstNull)
                {
                    BIH_DCA.Loan = (int)(sheet.Cells[i, 1] as Excel.Range).Value;
                    BIH_DCA.Client = (sheet.Cells[i, 2] as Excel.Range).Value;
                    BIH_DCA.DPD = (int)(sheet.Cells[i, 3] as Excel.Range).Value;
                    BIH_DCA.Bucket = (sheet.Cells[i, 4] as Excel.Range).Value;
                    BIH_DCA.Amount = (double)(sheet.Cells[i, 6] as Excel.Range).Value;
                    BIH_DCA.Percent = (double)(sheet.Cells[i, 7] as Excel.Range).Value;
                    BIH_DCA.Fee_amount = (double)(sheet.Cells[i, 8] as Excel.Range).Value;

                    try
                    {
                        ad_BIH_DCA_raw.InsertRow(BIH_DCA.Reestr_date.ToString("yyyy-MM-dd"), BIH_DCA.Loan, BIH_DCA.Client, BIH_DCA.DPD, BIH_DCA.Bucket, BIH_DCA.Debt_collector, BIH_DCA.Amount, BIH_DCA.Percent, BIH_DCA.Fee_amount);
                    }
                    catch (Exception exc)
                    {
                        COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                        logAdapter.InsertRow("cl_Parser_BIH", "BIH", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                    }

                    i++;
                }

                
            }
            catch (Exception exc)
            {
                COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_BIH", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
            }



            //Xml                                                           ----TO_DO
        }

        /*
        public void parse_BIH_SNAP(Excel.Application ex)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = sheet.get_Range("A1", last);
            lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;

            int i = 2; // Строка начала периода

            try
            {

                cl_BIH_SNAP BIH_SNAP = new cl_BIH_SNAP();

                string fileName = ex.Workbooks.Item[1].Name;
                fileName = fileName.Substring(fileName.IndexOf("_") + 1, 10); //.ToString("yyyy-MM-dd");

                DateTime reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Excel.Range).Value;
                BIH_SNAP.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);

                BIH_SNAP_rawTableAdapter ad_BIH_SNAP_raw = new BIH_SNAP_rawTableAdapter();
                ad_BIH_SNAP_raw.DeletePeriod(BIH_SNAP.Reestr_date.ToString("yyyy-MM-dd"));

                while (i <= lastUsedRow)
                {
                    BIH_SNAP.Loan = (int)(sheet.Cells[i, 1] as Excel.Range).Value;
                    BIH_SNAP.Current_status = (sheet.Cells[i, 2] as Excel.Range).Value;
                    BIH_SNAP.Loan_disbursement_date = DateTime.Parse((sheet.Cells[i, 3] as Excel.Range).Value);
                    BIH_SNAP.Product = (sheet.Cells[i, 4] as Excel.Range).Value;
                    BIH_SNAP.DPD = (int)(sheet.Cells[i, 5] as Excel.Range).Value;
                    BIH_SNAP.Historical_loan_status = (sheet.Cells[i, 6] as Excel.Range).Value;
                    BIH_SNAP.Principal_balance = (double)(sheet.Cells[i, 7] as Excel.Range).Value;
                    BIH_SNAP.Monthly_fee_balance = (double)(sheet.Cells[i, 8] as Excel.Range).Value;
                    BIH_SNAP.Guarantor_fee_balance = (double)(sheet.Cells[i, 9] as Excel.Range).Value;
                    BIH_SNAP.Penalty_fee_balance = (double)(sheet.Cells[i, 10] as Excel.Range).Value;
                    BIH_SNAP.Penalty_interest_balance = (double)(sheet.Cells[i, 11] as Excel.Range).Value;
                    BIH_SNAP.Interest_balance = (double)(sheet.Cells[i, 12] as Excel.Range).Value;

                    try
                    {
                        ad_BIH_SNAP_raw.InsertRow(BIH_SNAP.Reestr_date.ToString("yyyy-MM-dd"), BIH_SNAP.Loan, BIH_SNAP.Current_status, BIH_SNAP.Loan_disbursement_date.ToString("yyyy-MM-dd"),
                            BIH_SNAP.Product, BIH_SNAP.DPD, BIH_SNAP.Historical_loan_status, BIH_SNAP.Principal_balance, BIH_SNAP.Monthly_fee_balance, BIH_SNAP.Guarantor_fee_balance, BIH_SNAP.Penalty_fee_balance,
                            BIH_SNAP.Penalty_interest_balance, BIH_SNAP.Interest_balance);
                    }
                    catch (Exception exc)
                    {
                        COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                        logAdapter.InsertRow("cl_Parser", "BIH", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                    }

                    i++;
                }

                SP sp = new SP();
                sp.sp_BIH2_portfolio_snapshot();
                sp.sp_BIH_TOTAL_SNAP();

                Console.WriteLine("Loading is ready. " + (lastUsedRow - 1).ToString() + " rows were processed.");
            }
            catch (Exception exc)
            {
                COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser", "BIH", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
            }


            ex.Quit();

            Console.ReadKey();

            //Xml                                                           ----TO_DO

        }
        */

    }
}