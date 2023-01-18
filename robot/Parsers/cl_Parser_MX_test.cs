using Microsoft.Office.Interop.Excel;
using robot.DataSet1TableAdapters;
using robot.RiskTableAdapters;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static robot.Risk;

namespace robot.Parsers
{
    class cl_Parser_MX_test : cl_Parser
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
                    OpenFile(pathFile);
                    correctPath = 1;
                }
                catch
                {
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

        private void OpenFile(string pathFile)
        {
            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
            Application ex = new Application();
            Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл

            if (pathFile.Contains("cessions")) parse_MX_CESS(ex);
            if (pathFile.Contains("Data_exchance_format")) parse_MX_DCA(ex);
        }

        /*public void DcaPostProcessing()
        {

        }

        public void CessPostProcessing()
        {

        }*/

        private void parse_MX_CESS(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_MX", "parse_MX_CESS", "MX", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastUsedRow = last.Row; // Последняя строка в документе

            int firstNull = SearchFirstNullRow(sheet, lastUsedRow);
            //int firstNull = 12;

            int i = 2; // Строка начала периода

            try
            {
                string fileName = ex.Workbooks.Item[1].Name;

                fileName = "01." + fileName.Replace("cessions_", "").Replace(".xlsx", "").Substring(4, 2) + "." + fileName.Replace("cessions_", "").Replace(".xlsx", "").Substring(0, 4); //.ToString("yyyy-MM-dd");

                DateTime reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                                                                                                                 //MX_CESS.Reestr_date = reestr_date;       //current date
                MX_CESS_rawDataTable mx_cess_raw = new MX_CESS_rawDataTable();
                System.Data.DataTable mx_cess = new System.Data.DataTable();
                for (int j = 0; j < mx_cess_raw.Columns.Count; j++)
                    mx_cess.Columns.Add(mx_cess_raw.Columns[j].ColumnName, mx_cess_raw.Columns[j].DataType);

                while (i < firstNull)
                {
                    //MX_CESS_rawRow row = mx_cess.NewMX_CESS_rawRow();
                    System.Data.DataRow row = mx_cess.NewRow();

                    row["Reestr_date"] = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth

                    row["Loan_id"] = (double)(sheet.Cells[i, 1] as Range).Value;
                    row["Cession_date"] = (DateTime)(sheet.Cells[i, 2] as Range).Value;
                    row["Principal"] = (decimal)(sheet.Cells[i, 3] as Range).Value;
                    row["Interest"] = (sheet.Cells[i, 4] as Range).Text.ToString() != "#ЗНАЧ!" ? (decimal)(sheet.Cells[i, 4] as Range).Value : 0;
                    row["Fee"] = !(sheet.Cells[i, 5] as Range).Text.ToString().Contains("-") && !(sheet.Cells[i, 5] as Range).Text.ToString().Contains("$") && (sheet.Cells[i, 5] as Range).Text.ToString() != "" ? (decimal)(sheet.Cells[i, 5] as Range).Value : 0;
                    row["Penalty"] = !(sheet.Cells[i, 6] as Range).Text.ToString().Contains("-") && !(sheet.Cells[i, 6] as Range).Text.ToString().Contains("$") ? (decimal)(sheet.Cells[i, 6] as Range).Value : 0;
                    row["Otherdebt"] = !((sheet.Cells[i, 7] as Range).Text.ToString() == "") ? (decimal)(sheet.Cells[i, 7] as Range).Value : 0;
                    row["Price_amount"] = (decimal)(sheet.Cells[i, 8] as Range).Value;
                    row["Price_rate"] = (double)(sheet.Cells[i, 9] as Range).Value;
                    row["DPD"] = (int)(sheet.Cells[i, 11] as Range).Value;

                    //mx_cess.AddMX_CESS_rawRow(row);
                    mx_cess.Rows.Add(row);
                    mx_cess.AcceptChanges();

                    Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");

                    i++;
                }

                if (mx_cess.Rows.Count > 0)
                {
                    MX_CESS_rawTableAdapter ad_MX_CESS_raw = new MX_CESS_rawTableAdapter();
                    ad_MX_CESS_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"));


                    try
                    {
                        task = new cl_Tasks("exec Risk.dbo.sp_MX_CESS_raw @MX_CESS_raw = ", mx_cess);
                        //sprisk.sp_MX_CESS_raw(mx_cess);
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_MX", "parse_MX_CESS", "MX", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_descr: " + exc.Message);
                        ex.Quit();
                        //Console.ReadKey();

                        return;
                    }

                    report = "Loading is ready. " + (firstNull - 2).ToString() + " rows were processed.";
                    logAdapter.InsertRow("cl_Parser_MX", "parse_MX_CESS", "MX", DateTime.Now, true, report);
                    Console.WriteLine(report);

                    success = MX_Total_CESS_forming(reestr_date);
                    success += MX_Total_Snap_Cfield_forming();
                }
                else
                {
                    report = "File was empty. There is no one row.";
                    logAdapter.InsertRow("cl_Parser_MX", "parse_MX_CESS", "MX", DateTime.Now, false, report);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_descr: " + report);
                    ex.Quit();
                    Console.ReadKey();

                    return;
                }

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


            ex.Quit();

            if (success == 2)
            {
                //send_report = new cl_Send_Report("MX_CESS", 1);
                Console.WriteLine("Report was sent.");
            }

        }

        private void parse_MX_DCA(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_MX", "parse_MX_DCA", "MX", DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item("DCA");
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastUsedRow = last.Row; // Последняя строка в документе

            int firstNull = SearchFirstNullRow(sheet, lastUsedRow);

            int i = 2; // Строка начала периода
            
            try
            {
                string fileName = ex.Workbooks.Item[1].Name;

                string pattern = @"\D{4}\d{4}";
                Match result = Regex.Match(fileName, pattern);

                fileName = "01 " + result.ToString().Replace("_", " ");
                //"01 " + fileName.Replace("190725_Data_exchance_format_", "").Replace(".xlsx", "").Replace("_", " ");

                DateTime reestr_date = DateTime.Parse(fileName); //(DateTime)(sheet.Cells[i, 2] as Range).Value;
                reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth
                                                                                                                 //MX_DCA.Reestr_date = reestr_date;       //current date
                MX_DCA_rawDataTable mx_dca_raw = new MX_DCA_rawDataTable();
                System.Data.DataTable mx_dca = new System.Data.DataTable();
                for (int j = 0; j < mx_dca_raw.Columns.Count; j++)
                    mx_dca.Columns.Add(mx_dca_raw.Columns[j].ColumnName, mx_dca_raw.Columns[j].DataType);

                while (i < firstNull)
                {
                    System.Data.DataRow row = mx_dca.NewRow();

                    row["reestr_date"] = reestr_date;     //eomonth

                    row["loan_id"] = (sheet.Cells[i, 1] as Range).Value;
                    row["agreement_id"] = (sheet.Cells[i, 2] as Range).Value;
                    row["Payment_date"] = (DateTime)(sheet.Cells[i, 3] as Range).Value;
                    row["DCA_name"] = (sheet.Cells[i, 4] as Range).Value;
                    row["Payment_Amount"] = !(sheet.Cells[i, 5] as Range).Text.ToString().Contains("-") && !(sheet.Cells[i, 5] as Range).Text.ToString().Contains("$") && (sheet.Cells[i, 5] as Range).Text.ToString() != "" ? (decimal)(sheet.Cells[i, 5] as Range).Value : 0;
                    row["DCA_Comission_amount"] = !(sheet.Cells[i, 6] as Range).Text.ToString().Contains("-") && !(sheet.Cells[i, 6] as Range).Text.ToString().Contains("$") ? (decimal)(sheet.Cells[i, 6] as Range).Value : 0;

                    mx_dca.Rows.Add(row);
                    mx_dca.AcceptChanges();

                    Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");

                    i++;
                }

                if (mx_dca.Rows.Count > 0)
                {
                    MX_DCA_rawTableAdapter ad_MX_DCA_raw = new MX_DCA_rawTableAdapter();
                    ad_MX_DCA_raw.DeletePeriod(reestr_date.ToString("yyyy-MM-dd"));


                    try
                    {
                        task = new cl_Tasks("exec Risk.dbo.sp_MX_DCA_raw @MX_DCA_raw = ", mx_dca);
                        //sprisk.sp_MX_DCA_raw(mx_DCA);
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_MX", "parse_MX_DCA", "MX", DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_descr: " + exc.Message);
                        ex.Quit();
                        //Console.ReadKey();

                        return;
                    }

                    report = "Loading is ready. " + (firstNull - 2).ToString() + " rows were proDCAed.";
                    logAdapter.InsertRow("cl_Parser_MX", "parse_MX_DCA", "MX", DateTime.Now, true, report);
                    Console.WriteLine(report);

                    success = MX_Total_DCA_forming(reestr_date);
                }
                else
                {
                    report = "File was empty. There is no one row.";
                    logAdapter.InsertRow("cl_Parser_MX", "parse_MX_DCA", "MX", DateTime.Now, false, report);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_descr: " + report);
                    ex.Quit();
                    Console.ReadKey();

                    return;
                }

            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MX", "parse_MX_DCA", "MX", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                ex.Quit();
                Console.ReadKey();

                return;
            }

            
            ex.Quit();

            if (success == 1)
            {
                //send_report = new cl_Send_Report("MX_DCA", 1);
                Console.WriteLine("Report was sent.");
            }

        }

        private int MX_Total_DCA_forming(DateTime reestr_date)
        {
            try
            {
                task = new cl_Tasks("exec Risk.dbo.sp_MX_TOTAL_DCA @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //task_DCA.RunSynchronously();

                report = "TOTAL_DCA was formed successfully.";
                logAdapter.InsertRow("cl_Parser_MX", "MX_Total_DCA_forming", "MX", DateTime.Now, true, report);
                Console.WriteLine(report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MX", "MX_Total_DCA_forming", "MX", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }
        }

        private int MX_Total_CESS_forming(DateTime reestr_date)
        {
            object result;
            int indefinites = 0;

            //Task task_cess = new Task(() =>
            //{
            //    SPRisk sprisk = new SPRisk();
            //    result = sprisk.sp_MX_TOTAL_CESS(reestr_date);
            //    indefinites = int.Parse(result.ToString());
            //},
            //TaskCreationOptions.LongRunning);

            try
            {
                task = new cl_Tasks("exec Risk.dbo.sp_MX_TOTAL_CESS @date = '" + reestr_date.ToString("yyyy-MM-dd") + "'");
                //task_cess.RunSynchronously();

                report = indefinites == 1 ? "TOTAL_CESS was formed. Indefinite loan_ids were found." : "TOTAL_CESS was formed successfully.";
                logAdapter.InsertRow("cl_Parser_MX", "MX_Total_CESS_forming", "MX", DateTime.Now, true, report);
                Console.WriteLine(report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MX", "MX_Total_CESS_forming", "MX", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
            }
        }

        private int MX_Total_Snap_Cfield_forming()
        {
            try
            {
                task = new cl_Tasks("exec Risk.dbo.sp_MX_TOTAL_SNAP_CFIELD");
                //task_cess.RunSynchronously();

                report = "TOTAL_SNAP_CFIELD was formed successfully.";
                logAdapter.InsertRow("cl_Parser_MX", "MX_Total_Snap_Cfield_forming", "MX", DateTime.Now, true, report);
                Console.WriteLine(report);

                return 1;
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_MX", "MX_Total_Snap_Cfield_forming", "MX", DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_desc: " + exc.Message.ToString());

                return 0;
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

    }
}
