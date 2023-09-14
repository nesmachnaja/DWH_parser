using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace robot
{
    class cl_Field_mapping
    {
        int column_num = 0;
        int row_num = 0;

        public cl_Field_mapping(Worksheet sheet, string field_name, out int field_index)
        {
            field_index = 0;
            column_num = sheet.Columns.Count;
            row_num = 5; // sheet.Rows.Count;

            for (int i = 1; i <= row_num; i++)
            {
                for (int j = 1; j <= column_num; j++)
                {
                    if ((sheet.Cells[i, j] as Range).Value2 == null) break;
                    if ((sheet.Cells[i, j] as Range).Value.ToString().Replace("\n","").ToLower() == field_name.ToLower())
                    {
                        field_index = j;
                        return;
                    }
                }
            }

            //return field_index;
        }
        //public cl_Field_mapping(Worksheet sheet, string field_name, out int field_index)
        //{
        //    field_index = 0;
        //    column_num = sheet.Columns.Count;
        //    row_num = 5; // sheet.Rows.Count;

        //    for (int i = 1; i <= column_num; i++)
        //    {
        //        for (int j = 1; j <= row_num; j++)
        //        {
        //            if ((sheet.Cells[j, i] as Range).Value.ToString().ToLower() == field_name.ToLower())
        //            {
        //                field_index = i;
        //                break;
        //            }
        //        }
        //    }

        //    //return field_index;
        //}
    }
}
