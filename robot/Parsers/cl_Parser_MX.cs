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
    class cl_Parser_MX
    {
        //private int lastUsedRow;
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

        private void OpenFile(string pathFile)
        {            
            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
            Application ex = new Application();
            Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл

            if (pathFile.Contains("cessions")) parse_MX_CESS(ex);
        }

        private void parse_MX_CESS(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_MX", "parse_MX_CESS", "MX", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range range = sheet.get_Range("A1", last);
            int lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;

            int firstNull = SearchFirstNullRow(sheet, lastUsedRow);
            //int firstNull = 7000;

            cl_MX_CESS MX_CESS = new cl_MX_CESS();
            int i = 2; // Строка начала периода

            try
            {
                string fileName = ex.Workbooks.Item[1].Name;

                fileName = "01." + fileName.Replace("cessions_", "").Replace(".xlsx", "").Substring(4, 2) + "." + fileName.Replace("cessions_", "").Replace(".xlsx", "").Substring(0, 4); //.ToString("yyyy-MM-dd");

                DateTime reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                MX_CESS.Reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                                                                                                                         //MX_CESS.Reestr_date = reestr_date;       //current date

                //ex.Quit();

                MX_CESS_rawTableAdapter ad_MX_CESS_raw = new MX_CESS_rawTableAdapter();
                ad_MX_CESS_raw.DeletePeriod(MX_CESS.Reestr_date.ToString("yyyy-MM-dd"));

                while (i < firstNull)
                {
                    MX_CESS.Loan_id = (double)(sheet.Cells[i, 1] as Range).Value;
                    MX_CESS.Cession_date = (DateTime)(sheet.Cells[i, 2] as Range).Value;
                    MX_CESS.Principal = (decimal)(sheet.Cells[i, 3] as Range).Value;
                    MX_CESS.Interest = (decimal)(sheet.Cells[i, 4] as Range).Value;
                    MX_CESS.Fee = (decimal)(sheet.Cells[i, 5] as Range).Value;
                    MX_CESS.Penalty = (decimal)(sheet.Cells[i, 6] as Range).Value;
                    MX_CESS.Otherdebt = (decimal)(sheet.Cells[i, 7] as Range).Value;
                    MX_CESS.Price_amount = (decimal)(sheet.Cells[i, 8] as Range).Value;
                    MX_CESS.Price_rate = (double)(sheet.Cells[i, 9] as Range).Value;
                    MX_CESS.DPD = (int)(sheet.Cells[i, 10] as Range).Value;

                    try
                    {
                        ad_MX_CESS_raw.InsertRow(MX_CESS.Reestr_date.ToString("yyyy-MM-dd"), MX_CESS.Loan_id, MX_CESS.Cession_date, MX_CESS.Principal, MX_CESS.Interest,
                            MX_CESS.Fee, MX_CESS.Penalty, MX_CESS.Otherdebt, MX_CESS.Price_amount, MX_CESS.Price_rate, MX_CESS.DPD); //.ToString("yyyy-MM-dd"));
                        Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_MX", "parse_MX_CESS", "MX", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_descr: " + exc.Message);
                        ex.Quit();
                        Console.ReadKey();

                        return;
                    }

                    i++;
                }

                report = "Loading is ready. " + (firstNull - 1).ToString() + " rows were processed.";
                logAdapter.InsertRow("cl_Parser_MX", "parse_MX_CESS", "MX", DateTime.Now, true, report);
                Console.WriteLine(report);

                MX_Totoal_CESS_forming(MX_CESS.Reestr_date);

            }
            catch (Exception exc)
            {
                //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
                logAdapter.InsertRow("cl_Parser_MX", "parse_MX_CESS", "MX", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                ex.Quit();
                Console.ReadKey();

                return;
            }


            ex.Quit();

            /*
            Console.WriteLine("Do you want to transport snap to Risk? Y - Yes, N - No");
            string reply = Console.ReadKey().Key.ToString();

            
            if (reply.Equals("Y"))
            {
                TransportToRisk(MX_CESS.Reestr_date);
            }*/

            //report                                                           ----TO_DO

        }

        private void MX_Totoal_CESS_forming(DateTime reestr_date)
        {
            object result;
            int indefinites = 0;

            Task task_cess = new Task(() =>
            {
                SPRisk sprisk = new SPRisk();
                result = sprisk.sp_MX_TOTAL_CESS(reestr_date);
                indefinites = int.Parse(result.ToString());
            },
            TaskCreationOptions.LongRunning);

            try
            {
                task_cess.RunSynchronously();

                report = indefinites == 1 ? "TOTAL_CESS was formed. Indefinite loan_ids were found." : "TOTAL_CESS was formed successfully.";
                logAdapter.InsertRow("cl_Parser_MX", "parse_MX_CESS", "MX", DateTime.Now, true, report);
                Console.WriteLine(report);
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MX", "MX_Totoal_CESS_forming", "MX", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());
            }
        }

        /*
        private void TransportToRisk(DateTime reestr_date)
        {
            try
            {
                reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                SPRisk sprisk = new SPRisk();
                sprisk.sp_MX_TOTAL_CESS(reestr_date);

                Console.WriteLine("Cessions were transported to their destination on [Risk]");
                report = "Cessions were transported to their destination on [Risk]";
                logAdapter.InsertRow("cl_Parser_MX", "TransportToRisk", "MX", DateTime.Now, true, report);

            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MX", "TransportToRisk", "MX", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
            }

            Console.ReadKey();
        }
        */

        private static int SearchFirstNullRow(Worksheet sheet, int lastUsedRow)
        {
            int firstNull = 0;
            for (int firstEmpty = lastUsedRow + 1; firstEmpty > 1; firstEmpty--)
            {
                if (sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]) != 0)
                //&& sheet.Application.WorksheetFunction.CountA(sheet.Rows[firstEmpty]) == sheet.Application.WorksheetFunction.CountA(sheet.Rows[1]))
                {
                    firstNull = firstEmpty + 1;
                    break;
                }
            }

            return firstNull;
        }

    }
}
