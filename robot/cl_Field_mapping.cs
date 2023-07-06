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

        public cl_Field_mapping(Worksheet sheet, string field_name, out int field_index)
        {
            field_index = 0;
            column_num = sheet.Columns.Count;
            for (int i = 1; i <= column_num; i++)
            {
                if ((sheet.Cells[1, i] as Range).Value.ToString().ToLower() == field_name)
                {
                    field_index = i;
                    break;
                }
            }

            //return field_index;
        }
    }
}
