using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcel
{
    public static class ExcelExtensions
    {
        public static ExcelWorkbook GetExcelWorkbook(this DataTable DataTableToConvert)
        {
            ExcelWorkbook excelWorkbook = new ExcelWorkbook();

            excelWorkbook.AddSheet();
            excelWorkbook.Sheets[0].autoFilter = true;

            if (DataTableToConvert != null)
            {
                if (DataTableToConvert.Rows.Count > 0)
                {
                    foreach (DataRow dtRow in DataTableToConvert.Rows)
                    {
                        excelWorkbook.Sheets[0].AddRow();

                        //Adding the header row
                        if (excelWorkbook.Sheets[0].Rows.Count == 1)
                        {
                            foreach (DataColumn dtCol in DataTableToConvert.Columns)
                            {
                                excelWorkbook.Sheets[0].Rows.Last().AddColumn(new ExcelColumn() { Value = dtCol.ColumnName.ToString() });
                            }

                            excelWorkbook.Sheets[0].AddRow();
                        }
                        
                        foreach (DataColumn dtCol in DataTableToConvert.Columns)
                        {
                            excelWorkbook.Sheets[0].Rows.Last().AddColumn(new ExcelColumn() { Value = dtRow[dtCol.ColumnName].ToString() });
                        }
                    }
                }
            }

            return excelWorkbook;
        }
    }

}
