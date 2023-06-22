using Microsoft.Office.Interop.Excel;
using robot.DataSet1TableAdapters;
using robot.RiskTableAdapters;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static robot.DataSet1;

namespace robot.Parsers
{
    class cl_Parser_KZ_test : cl_Parser
    {
        string _country;
        string _databasename;

        public void StartParsing(string country, string path_file)
        {
            logAdapter = new COUNTRY_LogTableAdapter();
            int correctPath = 0;
            _country = country;

            while (correctPath == 0)
            {
                try
                {
                    pathFile = path_file;
                    OpenFile(pathFile);
                    correctPath = 1;
                }
                catch
                {
                    Console.WriteLine("Incorrect file path.");
                }
            }
        }

        /*private static string GetPath()
        {
            Console.WriteLine("Appoint file path: ");
            string pathFile = Console.ReadLine();
            return pathFile;
        }*/

        public void OpenFile(string pathFile)
        {
            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
            Application ex = new Application();
            Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл

            if (pathFile.Contains("port")) parse_KZ_SNAP(ex);
        }

        public void SnapPostProcessing()
        {
            success = TransportToCountryLevel();
            success += TotalSnapCFForming();

            success += TransportSnapToRisk();
            success += TransportSnapCFToRisk();

            
            if (success == 4)
            {
                send_report = new cl_Send_Report(_country.ToUpper() + "_SNAP", 1);
            }
        }

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

        private void parse_KZ_SNAP(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_KZ", "parse_KZ_SNAP", _country, DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            lastUsedRow = last.Row; // Последняя строка в документе

            int i = 2; // Строка начала периода

            int firstNull = SearchFirstNullRow(sheet, lastUsedRow);

            try
            {
                fileName = ex.Workbooks.Item[1].Name;

                KZ_SNAP_rawDataTable kz_snap_raw = new KZ_SNAP_rawDataTable();
                System.Data.DataTable kz_snap = new System.Data.DataTable();
                for (int j = 0; j < kz_snap_raw.Columns.Count; j++)
                    kz_snap.Columns.Add(kz_snap_raw.Columns[j].ColumnName, kz_snap_raw.Columns[j].DataType);

                brand = "Vivus";

                string pattern = @"(\d{2}\.\d{4})|(\d{2}\.\d{2})|(\d{4})";
                Match result = Regex.Match(fileName, pattern);

                fileName = "01." + result.ToString().Insert(2, ".");

                //fileName = "01." + fileName.Replace("portf_", "").Replace("kzfin_", "").Replace("vivus_", "").Replace(".xlsx", "").Insert(2, "."); //.Insert(5, "."); //.ToString("yyyy-MM-dd");

                reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                //KZ_SNAP.Reestr_date = reestr_date;       //current date

                while (i < firstNull)
                {
                    //KZ_SNAP_rawRow kz_snap_row = kz_snap.NewKZ_SNAP_rawRow();
                    System.Data.DataRow kz_snap_row = kz_snap.NewRow();

                    kz_snap_row["reestr_date"] = reestr_date;

                    kz_snap_row["ID_loan"] = (sheet.Cells[i, 1] as Range).Value.ToString();
                    kz_snap_row["Phone"] = (sheet.Cells[i, 2] as Range).Value.ToString();
                    kz_snap_row["Od"] = (double)(sheet.Cells[i, 3] as Range).Value;
                    kz_snap_row["Com"] = (sheet.Cells[i, 4] as Range).Value == null ? 0 : (double)(sheet.Cells[i, 4] as Range).Value;
                    kz_snap_row["Pen_balance"] = (double)(sheet.Cells[i, 5] as Range).Value;
                    kz_snap_row["Fee_balance"] = (double)(sheet.Cells[i, 6] as Range).Value;
                    kz_snap_row["Od_com"] = (double)(sheet.Cells[i, 7] as Range).Value;
                    kz_snap_row["Day_delay"] = (int)(sheet.Cells[i, 8] as Range).Value;
                    kz_snap_row["Date_start"] = DateTime.Parse((sheet.Cells[i, 9] as Range).Value);
                    kz_snap_row["ID_client"] = (sheet.Cells[i, 10] as Range).Value.ToString();
                    kz_snap_row["Interest"] = (sheet.Cells[i, 11] as Range).Value == null ? 0 : (double)(sheet.Cells[i, 11] as Range).Value;
                    kz_snap_row["Product"] = (sheet.Cells[i, 12] as Range).Value;
                    //kz_snap_row["Ces"] = (sheet.Cells[i, 12] as Range).Value;
                    kz_snap_row["Final_interest"] = (sheet.Cells[i, 14] as Range).Value == null ? 0 : (double)(sheet.Cells[i, 14] as Range).Value;
                    kz_snap_row["Prod"] = (sheet.Cells[i, 16] as Range).Value;
                    kz_snap_row["Status"] = (sheet.Cells[i, 15] as Range).Value;
                    kz_snap_row["CC"] = (sheet.Cells[i, 19] as Range).Value;

                    kz_snap_row["brand"] = brand;

                    //kz_snap.AddKZ_SNAP_rawRow(kz_snap_row);
                    kz_snap.Rows.Add(kz_snap_row);
                    kz_snap.AcceptChanges();

                    Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");

                    i++;

                }

                if (kz_snap.Rows.Count > 0)
                {
                    /*KZ_SNAP_rawTableAdapter ad_KZ_SNAP_raw = new KZ_SNAP_rawTableAdapter();
                    ad_KZ_SNAP_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"), brand);
                    ad_KZ_SNAP_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"), brand);*/

                    task = new cl_Tasks("delete from DWH_Risk.dbo." + _country + "_SNAP_raw where reestr_date = '" + reestr_date.ToString("yyyy-MM-dd") + "' and brand = '" + brand + "'");

                    /*Task task_kz_snap_raw = new Task(() =>
                    {
                        task = new cl_Tasks("exec DWH_Risk.dbo.sp_KZ_SNAP_raw @KZ_SNAP_raw = ", kz_snap);
                        //sp.sp_KZ_SNAP_raw(kz_snap);
                    },
                    TaskCreationOptions.LongRunning);*/

                    try
                    {
                        task = new cl_Tasks("exec DWH_Risk.dbo.sp_" + _country + "_SNAP_raw @" + _country + "_SNAP_raw = ", kz_snap);
                        //task_kz_snap_raw.RunSynchronously();
                        //sp.sp_KZ_SNAP_raw(kz_snap);
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_KZ", "parse_KZ_SNAP", _country, DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_descr: " + exc.Message);
                        ex.Quit();

                        return;
                    }


                    report = "Loading is ready. " + (firstNull - 2).ToString() + " rows were processed.";
                    logAdapter.InsertRow("cl_Parser_KZ", "parse_KZ_SNAP", _country, DateTime.Now, true, report);
                    Console.WriteLine(report);
                }
                else
                {
                    report = "File was empty. There is no one row.";
                    logAdapter.InsertRow("cl_Parser_KZ", "parse_KZ_SNAP", _country, DateTime.Now, false, report);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_descr: " + report);
                    ex.Quit();

                    return;
                }
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_KZ", "parse_KZ_SNAP", _country, DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                ex.Quit();

                return;
            }


            ex.Quit();

        }

        private int TotalSnapCFForming()
        {
            try
            {
                task = new cl_Tasks("exec " + _databasename + ".dbo.sp_TOTAL_SNAP_CFIELD");

                report = "[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                logAdapter.InsertRow("cl_Parser_KZ", "TotalSnapCFForming", _country, DateTime.Now, true, report);
                Console.WriteLine(report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_KZ", "TotalSnapCFForming", _country, DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }
        }

        private int TransportToCountryLevel()
        {
            _databasename = "Total_KZ";

            try
            {
                task = new cl_Tasks("exec " + _databasename + ".dbo.sp_TOTAL_SNAP @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");

                report = "[dbo].[TOTAL_SNAP] was transported to " + _databasename + ".";
                logAdapter.InsertRow("cl_Parser_KZ", "TransportToCountryLevel", _country, DateTime.Now, true, report);
                Console.WriteLine(report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_KZ", "TransportToCountryLevel", _country, DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }
        }

        private int TotalSnapForming()
        {
            try
            {
                task = new cl_Tasks("exec DWH_Risk.dbo.sp_" + _country + "_TOTAL_SNAP @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //sp.sp_KZ_TOTAL_SNAP(reestr_date);

                report = "[dbo].[TOTAL_SNAP] was formed.";
                logAdapter.InsertRow("cl_Parser_KZ", "TotalSnapForming", _country, DateTime.Now, true, report);
                Console.WriteLine(report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_KZ", "TotalSnapForming", _country, DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }
        }

        private int TransportSnapToRisk()
        {
            try
            {
                task = new cl_Tasks("exec Risk.dbo.sp_KZ_TOTAL_SNAP @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");

                //Console.WriteLine("Snap was transported to [Risk].[dbo].[TOTAL_SNAP].");
                report = "Snap was transported to [Risk].[dbo].[TOTAL_SNAP].";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_KZ", "TransportSnapToRisk", _country, DateTime.Now, true, report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_KZ", "TransportSnapToRisk", _country, DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }

        }

        private int TransportSnapCFToRisk()
        {
            try
            {
                task = new cl_Tasks("exec Risk.dbo.sp_KZ_TOTAL_SNAP_CFIELD @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");

                report = "[Risk].[dbo].[TOTAL_SNAP_CFIELD] was formed.";
                Console.WriteLine(report);
                logAdapter.InsertRow("cl_Parser_KZ", "TransportSnapCFToRisk", _country, DateTime.Now, true, report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_KZ", "TransportSnapCFToRisk", _country, DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }

        }
    }
}
