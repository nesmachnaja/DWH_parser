using Microsoft.Office.Interop.Excel;
using robot.DataSet1TableAdapters;
using robot.Total_DBTableAdapters;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static robot.Total_DB;

namespace robot.Parsers
{
    class cl_Parser_budget : cl_Parser
    {
        string _country;
        string _databasename;

        public void StartParsing(string path_file)
        {
            logAdapter = new COUNTRY_LogTableAdapter();
            int correctPath = 0;

            string pattern = @"\b\w*_";
            Match result = Regex.Match(path_file, pattern);

            _country = result.ToString().Replace("_budg_","");

            while (correctPath == 0)
            {
                try
                {
                    pathFile = path_file;
                    OpenFile(pathFile);
                    correctPath = 1;
                }
                catch (Exception exc)
                {
                    Console.WriteLine("Incorrect file path.");
                    Console.WriteLine(exc.Message);
                }
            }
        }

        public void OpenFile(string pathFile)
        {
            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
            Application ex = new Application();
            Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открываем файл

            parse_budget(ex);
        }

        public void SendReporting()
        {
            if (success == 1)
            {
                send_report = new cl_Send_Report("budget_" + _country, 1);
            }
        }

        private void parse_budget(Application ex)
        {
            report = "Loading started.";
            logAdapter.InsertRow("cl_Parser_budget", "parse_budget", _country, DateTime.Now, true, report);

            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(1); // берем первый лист;
            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range range = sheet.get_Range("A1", last);
            int lastUsedRow = last.Row; // Последняя строка в документе
            int lastUsedColumn = last.Column;

            int firstNull = SearchFirstNullRow(sheet, lastUsedRow);

            budget_rawDataTable budget_raw = new budget_rawDataTable();
            System.Data.DataTable budget = new System.Data.DataTable();
            for (int j = 1; j < budget_raw.Columns.Count; j++)
                budget.Columns.Add(budget_raw.Columns[j].ColumnName, budget_raw.Columns[j].DataType);

            int i = 2; // Строка начала периода

            new cl_Field_mapping(sheet, "country", out int country);
            new cl_Field_mapping(sheet, "product", out int product);
            new cl_Field_mapping(sheet, "channel", out int channel);
            new cl_Field_mapping(sheet, "cycle_type", out int cycle_type);
            new cl_Field_mapping(sheet, "cost_(EUR)", out int cost);
            new cl_Field_mapping(sheet, "last day of the week", out int last_day_week);
            new cl_Field_mapping(sheet, "last day of the month", out int last_day_month);

            try
            {
                fileName = ex.Workbooks.Item[1].Name;

                string pattern = @"\d+";
                Match result = Regex.Match(fileName, pattern);

                reestr_date = DateTime.Parse(result.Value.ToString().Insert(2,".").Insert(5,"."));
                //reestr_date = new DateTime(reestr_date.Year, reestr_date.Month, 1).AddMonths(1).AddDays(-1);     //eomonth

                while (i < firstNull)
                {
                    System.Data.DataRow budget_row = budget.NewRow();
                    //budget_rawRow budget_row = budget.Newbudget_rawRow();

                    //budget_row["reestr_date"] = reestr_date;

                    budget_row["country"] = (sheet.Cells[i, country] as Range).Value.ToString();
                    budget_row["product"] = (sheet.Cells[i, product] as Range).Value.ToString();
                    budget_row["channel"] = (sheet.Cells[i, channel] as Range).Value.ToString();
                    budget_row["cycle_type"] = (sheet.Cells[i, cycle_type] as Range).Value.ToString();
                    budget_row["cost"] = (double?)(sheet.Cells[i, cost] as Range).Value == null ? 0 : (double?)(sheet.Cells[i, 5] as Range).Value;
                    budget_row["last_day_week"] = (DateTime)(sheet.Cells[i, last_day_week] as Range).Value;
                    budget_row["last_day_month"] = (DateTime)(sheet.Cells[i, last_day_month] as Range).Value;
                    

                    //budget.Addbudget_rawRow(budget_row);
                    budget.Rows.Add(budget_row);
                    budget.AcceptChanges();

                    Console.WriteLine((i - 1).ToString() + "/" + (firstNull - 2).ToString() + " row uploaded");

                    i++;
                }

                if (budget.Rows.Count > 0)
                {
                    budget_rawTableAdapter ad_budget_raw = new budget_rawTableAdapter();
                    ad_budget_raw.DeleteCountryPeriod(_country, reestr_date.ToString("yyyy-MM-dd"));


                    try
                    {
                        task = new cl_Tasks("exec Total_DB.dbo.sp_budget_raw @budget_raw = ", budget);
                        //sp.sp_budget_raw(budget);

                        if (task.success != 1) throw new Exception();
                        
                        success = 1;

                        SendReporting();
                    }
                    catch (Exception exc)
                    {
                        logAdapter.InsertRow("cl_Parser_budget", "parse_budget", _country, DateTime.Now, false, exc.Message);
                        Console.WriteLine("Error");
                        Console.WriteLine("Error_descr: " + exc.Message);
                        ex.Quit();

                        return;
                    }
                }
                else
                {
                    report = "File was empty. There is no one row.";
                    logAdapter.InsertRow("cl_Parser_budget", "parse_budget", _country, DateTime.Now, false, report);
                    Console.WriteLine("Error");
                    Console.WriteLine("Error_descr: " + report);
                    ex.Quit();

                    return;
                }
            }
            catch (Exception exc)
            {
                logAdapter.InsertRow("cl_Parser_budget", "parse_budget", _country, DateTime.Now, false, exc.Message);
                Console.WriteLine("Error");
                Console.WriteLine("Error_descr: " + exc.Message);
                ex.Quit();

                return;
            }


            ex.Quit();

            report = "Loading is ready. " + (firstNull - 1).ToString() + " rows were processed.";
            logAdapter.InsertRow("cl_Parser_budget", "parse_budget", _country, DateTime.Now, true, report);
            Console.WriteLine(report);

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
