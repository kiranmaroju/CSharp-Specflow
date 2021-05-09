using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
namespace Ledger_AutomationTesting.ExcelUtilities
{
    class ExcelLib
    {
        public string filename = @"C:\Users\Ramesh\Downloads\Learn_Specflow-master\Fujitsu_Assignment\Learn_Specflow\TestData\TestData.xlsx";
        public void writeToExcel(string sheet, int row, int column, string val)
        {
            Excel.Application excel_app = new Excel.Application();
            Excel.Workbook workbook = excel_app.Workbooks.Open(filename);
            excel_app.Visible = false;
            Excel.Worksheet sh = workbook.Sheets[sheet];
            sh.Cells[row, column] =val;
            workbook.Save();
            excel_app.Quit();
        }

        public string getValueFromDataSheet(string sheet, int row, int column)
        {
            
            Excel.Application excel_app = new Excel.Application();
            Excel.Workbook workbook = excel_app.Workbooks.Open(filename);
            excel_app.Visible = false;
            Excel.Worksheet sh = workbook.Sheets[sheet];
            string cellValue=sh.Cells[row, column];
            //workbook.Save();
            excel_app.Quit();
            return cellValue;
        }


        public string getValueFromExcel(string sheet, int row, int column)
        {
            Excel.Application excel_app = new Excel.Application();

            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Ramesh\Downloads\Learn_Specflow-master\Fujitsu_Assignment\Learn_Specflow\TestData\TestData.xlsx;Extended Properties=Excel 8.0;";

            var dataTable = new DataTable();
            OleDbConnection con = new OleDbConnection(connectionString);
            string query = string.Format("SELECT * FROM ["+ sheet + "$]");
            con.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
            adapter.Fill(dataTable);
            con.Close();
            return dataTable.Rows[row-1][column-1].ToString();
            
        }
        public DataTable ExcelToDataTable(string sheet)
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Ramesh\Downloads\Learn_Specflow-master\Fujitsu_Assignment\Learn_Specflow\TestData\TestData.xlsx;Extended Properties=Excel 8.0;";
            var dataTable = new DataTable();
            OleDbConnection con = new OleDbConnection(connectionString);
            string query = string.Format("SELECT * FROM ["+ sheet + "$]");
            con.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
            adapter.Fill(dataTable);
            con.Close();
            return dataTable;
        }

        List<DataCollection> dataCol = new List<DataCollection>();
        
        //This function populate the DataCollection class
        public void PopulateInCollection()
        {
            DataTable table = ExcelToDataTable("");

            for(int row = 1; row <= table.Rows.Count; row++)
            {
                for(int col = 0; col<table.Columns.Count; col++)
                {
                    DataCollection dtTable = new DataCollection()
                    {
                        rowNumber = row,
                        colName = table.Columns[col].ColumnName,
                        colValue = table.Rows[row - 1][col].ToString()
                    };
                    dataCol.Add(dtTable);
                }
            }
        }

        //This method can be directly use to read the data from Excel, just pass the row no and cloumn name Example (1, "UserName")
        public string ReadData(int RowNumer, string columnName)
        {
            try
            {
                string data = (from colData in dataCol
                               where colData.colName == columnName && colData.rowNumber == RowNumer
                               select colData.colValue).SingleOrDefault();

                return data.ToString();
            }
            catch (Exception e)
            {
                return null;
            }
        }


       

    }

    public class DataCollection
    {
        public int rowNumber { get; set; }
        public string colName { get; set; }
        public string colValue { get; set; }
    }
}
