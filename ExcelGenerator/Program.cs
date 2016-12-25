using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenExcel;
using System.Data;

namespace ExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //GenerateExcelFromData();

                GenerateExcelFromDataTable();

                Console.WriteLine("\n\n\t\tExcel file generated successfully\n\n\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }            
        }

        /// <summary>
        /// Method to generate excel from data
        /// </summary>
        private static void GenerateExcelFromData()
        {
            ExcelWorkbook newWorkbook = new ExcelWorkbook("ExcelFromData");

            newWorkbook.AddSheet("New Sheet");
            newWorkbook.Sheets[0].AddRow();
            newWorkbook.Sheets[0].AddRow();

            ExcelColumn col1 = new ExcelColumn();
            col1.Value = "First Name";

            ExcelColumn col2 = new ExcelColumn();
            col2.Value = "Last Name";

            ExcelColumn col3 = new ExcelColumn();
            col3.Value = "Anshul";

            ExcelColumn col4 = new ExcelColumn();
            col4.Value = "Chandra";

            newWorkbook.Sheets[0].Rows[0].AddColumn(col1);
            newWorkbook.Sheets[0].Rows[0].AddColumn(col2);
            newWorkbook.Sheets[0].Rows[1].AddColumn(col3);
            newWorkbook.Sheets[0].Rows[1].AddColumn(col4);

            newWorkbook.Download();
        }

        /// <summary>
        /// Method to generate excel from DataTable
        /// </summary>
        private static void GenerateExcelFromDataTable()
        {
            DataTable dtSample = new DataTable("ExcelFromDataTable");

            dtSample.Columns.Add("ID", typeof(string));
            dtSample.Columns.Add("Name", typeof(string));

            dtSample.Rows.Add();
            dtSample.Rows.Add();
            dtSample.Rows.Add();
            dtSample.Rows.Add();

            dtSample.Rows[0]["ID"] = "1";
            dtSample.Rows[0]["Name"] = "Anshul";
            dtSample.Rows[1]["ID"] = "2";
            dtSample.Rows[1]["Name"] = "Abhinav";
            dtSample.Rows[2]["ID"] = "3";
            dtSample.Rows[2]["Name"] = "Ankur";
            dtSample.Rows[3]["ID"] = "4";
            dtSample.Rows[3]["Name"] = "Navanshu";

            ExcelWorkbook newWorkbook = dtSample.GetExcelWorkbook();

            newWorkbook.Download();
        }
    }
}
