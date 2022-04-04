//using robot.DataSet1TableAdapters;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Microsoft.Office.Interop.Excel;
//using robot.Structures;

//namespace robot.Parsers
//{
//    class cl_Parser_MD
//    {
//        COUNTRY_LogTableAdapter logAdapter;
//        cl_MD_DCA MD_DCA = new cl_MD_DCA();

//        public void OpenFile()
//        {
//            logAdapter = new COUNTRY_LogTableAdapter();

//            string pathFile = @"C:\Users\Людмила\source\repos\robot\external_collection_02.2022.xlsx"; // Путь к файлу отчета
//            //static string pathFile = @"C:\Users\Людмила\source\repos\robot\DCA.xlsx"; // Путь к файлу отчета
//            string fullPath = Path.GetFullPath(pathFile); // Заплатка для корректности прав
//            Application ex = new Application();
//            Workbook workBook = ex.Workbooks.Open(fullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
//                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
//                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
//                Type.Missing, Type.Missing); //открываем файл

//            if (pathFile.Contains("Plati")) parse_MD_DCA(ex);
//            //if (pathFile.Contains("snapshot")) parse_MD_SNAP(ex);
//        }


//        public void parse_MD_DCA(Application ex)
//        {
//            string report = "Loading started.";
//            logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA", "MD", DateTime.Now, true, report);

//            int lastUsedRow = 0;
//            string fileName = ex.Workbooks.Item[1].Name;

//            int startIndex = fileName.LastIndexOf("_") + 1;
//            fileName = "01." + fileName.Substring(startIndex, fileName.Length - startIndex).Replace(".xlsx", "");
//            MD_DCA.Reestr_date = DateTime.Parse(fileName).AddMonths(1).AddDays(-1);

            
//            Worksheet sheet = (Worksheet)ex.Worksheets.get_Item(2); // берем первый лист;
//            //Console.WriteLine("Sheet #2");
//            parse_MD_DCA_current_sheet(sheet);
            

//            try
//            {
//                SP sp = new SP();
//                sp.sp_MD2_DCA(MD_DCA.Reestr_date);
//                sp.sp_MD_TOTAL_DCA(MD_DCA.Reestr_date);
//                Console.WriteLine("Loading is ready. " + lastUsedRow.ToString() + " rows were processed.");
//            }
//            catch (Exception exc)
//            {
//                //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
//                logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA", "MD", DateTime.Now, false, exc.Message);
//                Console.WriteLine("Error");
//            }

//            ex.Quit();

//            report = "Loading is ready. " + lastUsedRow.ToString() + " rows were processed.";
//            logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA", "MD", DateTime.Now, true, report);

//            Console.ReadKey();

//        }

//        private void parse_MD_DCA_current_sheet(Worksheet sheet)
//        {
//            Range last = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
//            Range range = sheet.get_Range("A1", last);
//            int lastUsedRow = last.Row;
//            int lastUsedColumn = last.Column;

//            int i = 2; // Строка начала периода

//            try
//            {
//                MD_DCA_rawTableAdapter ad_MD_DCA_raw = new MD_DCA_rawTableAdapter();
//                ad_MD_DCA_raw.DeletePeriod(MD_DCA.Reestr_date.ToString("yyyy-MM-dd"));

//                while (i < lastUsedRow)
//                {
//                    MD_DCA.Collection_company = (sheet.Cells[i, 1] as Range).Value;
//                    MD_DCA.Payment_month = DateTime.Parse((sheet.Cells[i, 2] as Range).Value);
//                    MD_DCA.Debtor = (sheet.Cells[i, 3] as Range).Value;
//                    MD_DCA.IDNP_debitorului = (sheet.Cells[i, 4] as Range).Value;
//                    MD_DCA.Contract = (int)(sheet.Cells[i, 5] as Range).Value;
//                    MD_DCA.Total_paid = (double)(sheet.Cells[i, 6] as Range).Value;
//                    MD_DCA.Fee = (double)(sheet.Cells[i, 7] as Range).Value;
//                    MD_DCA.Fee_including_VAT = (double)(sheet.Cells[i, 7] as Range).Value;
//                    MD_DCA.Types = (sheet.Cells[i, 7] as Range).Value;
//                    MD_DCA.Payment_date = DateTime.Parse((sheet.Cells[i, 7] as Range).Value);

//                    try
//                    {
//                        ad_MD_DCA_raw.InsertRow(MD_DCA.Reestr_date.ToString("yyyy-MM-dd"), MD_DCA.Collection_company, MD_DCA.Payment_month.ToString("yyyy-MM-dd"), MD_DCA.Debtor, MD_DCA.IDNP_debitorului, MD_DCA.Contract, MD_DCA.Total_paid, MD_DCA.Fee, MD_DCA.Fee_including_VAT,
//                            MD_DCA.Types, MD_DCA.Payment_date.ToString("yyyy-MM-dd"));
//                        Console.WriteLine((i - 1).ToString() + "/" + (lastUsedRow - 1).ToString() + " row uploaded");
//                    }
//                    catch (Exception exc)
//                    {
//                        //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
//                        logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA_current_sheet", "MD", DateTime.Now, false, exc.Message);
//                        Console.WriteLine("Error");
//                    }

//                    i++;
//                }


//            }
//            catch (Exception exc)
//            {
//                //COUNTRY_LogTableAdapter logAdapter = new COUNTRY_LogTableAdapter();
//                logAdapter.InsertRow("cl_Parser_MD", "parse_MD_DCA_current_sheet", "MD", DateTime.Now, false, exc.Message);
//                Console.WriteLine("Error");
//            }



//            //report                                                           ----TO_DO
//        }

//    }
//}
