using robot.DataSet1TableAdapters;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using robot.Structures;
using robot.RiskTableAdapters;

namespace robot.Parsers
{
    class cl_Parser_MD
    {
        COUNTRY_LogTableAdapter logAdapter;
        int lastUsedRow;
        string report;

        public void OpenFile()
        {
            logAdapter = new COUNTRY_LogTableAdapter();

            string pathFile = @"C:\Users\Людмила\source\repos\robot\Plati colectate totalizator generalizat INCASO PFBC si Agerlex_08.02.22_v2.xlsx"; // Путь к файлу отчета
            //static string pathFile = @"C:\Users\Людмила\source\repos\robot\DCA.xlsx"; // Путь к файлу отчета
            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
            Application ex = new Application();
            Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл

            if (pathFile.Contains("Plati")) parse_MD_DCA(ex);
            if (pathFile.Contains("SNAP") || pathFile.Contains("WO")) parse_MD_SNAP(ex);
        }


        public void parse_MD_DCA(Application ex)
        {
            string report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA", "MD", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(2); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range range = sheet.get_Range("A1", last);
            int lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;
            cl_MD_DCA MD_DCA = new cl_MD_DCA();


            int i = lastUsedRow; // Строка начала периода

            try
            {
                DateTime reestr_date = (DateTime)(sheet.Cells[i, 2] as Range).Value;
                MD_DCA.Reestr_date = reestr_date;
                //MD_DCA.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);

                MD_DCA_rawTableAdapter ad_MD_DCA_raw = new MD_DCA_rawTableAdapter();
                ad_MD_DCA_raw.DeletePeriod(MD_DCA.Reestr_date.ToString("yyyy-MM-dd"));

                while (i > 0)
                {
                    MD_DCA.Collection_company = (sheet.Cells[i, 1] as Range).Value;
                    MD_DCA.Payment_month = (DateTime)(sheet.Cells[i, 2] as Range).Value;
                    MD_DCA.Debtor = (sheet.Cells[i, 3] as Range).Value;
                    MD_DCA.IDNP_debitorului = (sheet.Cells[i, 4] as Range).Value;
                    MD_DCA.Contract = (sheet.Cells[i, 5] as Range).Value;
                    MD_DCA.Total_paid = (double)(sheet.Cells[i, 6] as Range).Value;
                    MD_DCA.Fee = (double)(sheet.Cells[i, 7] as Range).Value;
                    MD_DCA.Fee_including_VAT = (double)(sheet.Cells[i, 8] as Range).Value;
                    MD_DCA.Types = (sheet.Cells[i, 9] as Range).Value;
                    MD_DCA.Payment_date = DateTime.Parse((sheet.Cells[i, 10] as Range).Value.ToString().Replace("0:00:00",""));

                    if (MD_DCA.Payment_month == MD_DCA.Reestr_date)
                    {
                        try
                        {
                            ad_MD_DCA_raw.InsertRow(MD_DCA.Reestr_date.ToString("yyyy-MM-dd"), MD_DCA.Collection_company, MD_DCA.Payment_month.ToString("yyyy-MM-dd"), MD_DCA.Debtor, MD_DCA.IDNP_debitorului, MD_DCA.Contract, MD_DCA.Total_paid, MD_DCA.Fee,
                            MD_DCA.Fee_including_VAT, MD_DCA.Types, MD_DCA.Payment_date.ToString("yyyy-MM-dd"));
                            Console.WriteLine((lastUsedRow - i + 1).ToString() + "/" + (lastUsedRow - 1).ToString() + " row uploaded");
                        }
                        catch (Exception exc)
                        {
                            logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA", "MD", DateTime.Now, false, exc.Message);
                            Console.WriteLine("Error");
                            ex.Quit();
                        }

                        i--;
                    }
                    else 
                    {
                        Console.WriteLine("The other rows are marked by another Payment_month");
                        ex.Quit();
                        Console.ReadKey();
                        break;
                    }
                }

                SP sp = new SP();
                sp.sp_MD2_DCA(MD_DCA.Reestr_date);
                sp.sp_MD_TOTAL_DCA();
                Console.WriteLine("Loading is ready. " + (lastUsedRow - i).ToString() + " rows were processed.");
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                ex.Quit();
                Console.ReadKey();
                return;
            }

            //ex.Quit();

            //SP sp = new SP();
            //sp.sp_MD_TOTAL_DCA(MD_DCA.Reestr_date);

            report = "Loading is ready. " + (lastUsedRow - i).ToString() + " rows were processed.";
            logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA", "MD", DateTime.Now, true, report);

            Console.WriteLine("Do you want to transport DCA to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                TransportDCAToRisk(MD_DCA.Reestr_date);
            }

            //Console.ReadKey();

            //Xml                                                           ----TO_DO

        }


        public void parse_MD_SNAP(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_MD", "parse_MD_SNAP", "MD", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range range = sheet.get_Range("A1", last);
            lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;
            cl_MD_SNAP MD_SNAP = new cl_MD_SNAP();

            int firstNull = SearchFirstNullRow(sheet, lastUsedRow);

            int i = 3; // Строка начала периода
            int startPosition = i - 1; // Строка начала периода

            try
            {
                string fileName = ex.Workbooks.Item[1].Name;
                fileName = fileName.Replace("Moldova_SNAP ","").Replace("Moldova_WO ","").Replace("Moldova_WO_accumulated_","").Replace(".xlsx","").Replace("_"," "); //.ToString("yyyy-MM-dd");

                //ex.Quit();

                DateTime reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                //MD_SNAP.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                MD_SNAP.Reestr_date = reestr_date;       //current date
                MD_SNAP.Snapdate = MD_SNAP.Reestr_date;       //current date

                MD_SNAP.Source_type = ex.Workbooks.Item[1].Name.Replace(".xlsx", "");

                MD_SNAP_rawTableAdapter ad_MD_SNAP_raw = new MD_SNAP_rawTableAdapter();
                ad_MD_SNAP_raw.DeletePeriod(MD_SNAP.Reestr_date.ToString("yyyy-MM-dd"), MD_SNAP.Source_type);

                while (i < firstNull)
                {
                    MD_SNAP.Account_ID = (sheet.Cells[i, 1] as Range).Value.ToString();
                    MD_SNAP.Loan_amount = (double)(sheet.Cells[i, 4] as Range).Value;
                    MD_SNAP.DPD = (int)(sheet.Cells[i, 23] as Range).Value;
                    MD_SNAP.Principal_balance = (double)(sheet.Cells[i, 7] as Range).Value;
                    MD_SNAP.Principal = (double)(sheet.Cells[i, 8] as Range).Value;
                    MD_SNAP.Origination_fee = (double)(sheet.Cells[i, 9] as Range).Value;
                    MD_SNAP.Origination_fee_IL = (double)(sheet.Cells[i, 10] as Range).Value;
                    MD_SNAP.Interest_balance_for_provisions = (double)(sheet.Cells[i, 11] as Range).Value;

                    try
                    {
                        ad_MD_SNAP_raw.InsertRow(MD_SNAP.Reestr_date.ToString("yyyy-MM-dd"), MD_SNAP.Snapdate.ToString("yyyy-MM-dd"), MD_SNAP.Account_ID, MD_SNAP.Loan_amount, MD_SNAP.DPD,
                            MD_SNAP.Principal_balance, MD_SNAP.Principal, MD_SNAP.Origination_fee, MD_SNAP.Origination_fee_IL, MD_SNAP.Interest_balance_for_provisions, MD_SNAP.Source_type);
                        Console.WriteLine((i - startPosition).ToString() + "/" + (firstNull - startPosition - 1).ToString() + " row uploaded");
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_MD", "parse_MD_SNAP", "MD", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_descr: " + exc.Message.ToString());
                        ex.Quit();
                    }

                    i++;
                }

                ad_MD_SNAP_raw.UpdateInitialsAndClients();

                //SP sp = new SP();
                //sp.sp_MD2_portfolio_snapshot(MD_SNAP.Reestr_date);
                //sp.sp_MD_TOTAL_SNAP(MD_SNAP.Reestr_date);

                Console.WriteLine("Loading is ready. " + (firstNull - startPosition - 1).ToString() + " rows were processed.");

                report = "Loading is ready. " + (firstNull - startPosition - 1).ToString() + " rows were processed.";
                logAdapter.InsertRow("cl_Parser_MD", "parse_MD_SNAP", "MD", DateTime.Now, true, report);

            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "parse_MD_SNAP", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message.ToString());
                ex.Quit();
            }


            ex.Quit();

            Console.WriteLine("Do you want to transport snap to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();


            if (reply.Equals("Y"))
            {
                TransportSnapToRisk(MD_SNAP.Snapdate);
            }
            //report                                                           ----TO_DO

        }

        private void TransportSnapToRisk(DateTime snapdate)
        {
            try
            {
                SPRisk sprisk = new SPRisk();
                sprisk.sp_MD2_portfolio_snapshot(snapdate);
                Console.WriteLine("Snap was transported to [Risk].[dbo].[MD2_portfolio_snapshot]");
                report = "Snap was transported to [Risk].[dbo].[MD2_portfolio_snapshot]";
                logAdapter.InsertRow("cl_Parser_MD", "TransportSnapToRisk", "MD", DateTime.Now, true, report);

                //report
                sprisk.sp_MD3_portfolio_snapshot(snapdate);
                Console.WriteLine("IL-block was calculated in [Risk].[dbo].[MD3_portfolio_snapshot]");
                report = "IL-block was calculated in [Risk].[dbo].[MD3_portfolio_snapshot]";
                logAdapter.InsertRow("cl_Parser_MD", "TransportSnapToRisk", "MD", DateTime.Now, true, report);
                
                //report
                sprisk.sp_MD_TOTAL_SNAP();
                Console.WriteLine("[Risk].[dbo].[TOTAL_SNAP] was formed.");
                report = "[Risk].[dbo].[TOTAL_SNAP] was formed.";
                logAdapter.InsertRow("cl_Parser_MD", "TransportSnapToRisk", "MD", DateTime.Now, true, report);
                
                //report
                sprisk.sp_MD_TOTAL_SNAP_CFIELD();
                Console.WriteLine("[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.");
                report = "[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                logAdapter.InsertRow("cl_Parser_MD", "TransportSnapToRisk", "MD", DateTime.Now, true, report);

                //report into log
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "TransportSnapToRisk", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }

            Console.ReadKey();

        }

        private void TransportDCAToRisk(DateTime t_date)
        {
            try
            {
                SPRisk sprisk = new SPRisk();
                sprisk.sp_MD_TOTAL_DCA(t_date);
                Console.WriteLine("DCA was transported to [Risk].[dbo].[MD2_DCA], [Risk].[dbo].[TOTAL_DCA]");
                report = "DCA was transported to [Risk].[dbo].[MD2_DCA], [Risk].[dbo].[TOTAL_DCA]";
                logAdapter.InsertRow("cl_Parser_MD", "TransportDCAToRisk", "MD", DateTime.Now, true, report);
                //report into log
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MD", "TransportDCAToRisk", "MD", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }

            Console.ReadKey();

        }

        private static int SearchFirstNullRow(Worksheet sheet, int lastUsedRow)
        {
            int firstNull = 0;
            for (int firstEmpty = lastUsedRow; firstEmpty > 0; firstEmpty--)
            {
                if (sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]) != 0)
                {
                    firstNull = firstEmpty + 1;
                    break;
                }
            }

            return firstNull;
        }

    }
}
