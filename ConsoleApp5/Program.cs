using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace ConsoleApp5
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string excelFilePath = @"C:\Users\turbim\Desktop\MamulDepo.xlsx";
            string connectionString = "Data Source=192.168.8.100;Initial Catalog=CILSAN; User Id=turbim;Password=Turbim27; Integrated Security=True";

            // Excel'den verileri oku
            System.Data.DataTable dataTable = ReadExcelData(excelFilePath);

            // SQL Server'a verileri aktar
            WriteToSqlServer(dataTable, connectionString);

            Console.WriteLine("Veri aktarımı tamamlandı.");
            Console.ReadLine();
        }

        static System.Data.DataTable ReadExcelData(string filePath)
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1];
            Range range = worksheet.UsedRange;

            System.Data.DataTable dataTable = new System.Data.DataTable();

            for (int row = 1; row <= range.Rows.Count; row++)
            {
                if (row == 1)
                {
                    // Sütun başlıklarını ekleyin
                    for (int col = 1; col <= range.Columns.Count; col++)
                    {
                        dataTable.Columns.Add((string)(range.Cells[row, col] as Range).Value2);
                    }
                }
                else
                {
                    // Verileri ekleyin
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= range.Columns.Count; col++)
                    {
                        dataRow[col - 1] = (range.Cells[row, col] as Range).Value2;
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }

            workbook.Close(false);
            excelApp.Quit();

            return dataTable;
        }

        static void WriteToSqlServer(System.Data.DataTable dataTable, string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                {
                    bulkCopy.DestinationTableName = "CRD_Items";

                    // Açıkça sütun eşleştirmelerini ekle
                    bulkCopy.ColumnMappings.Add("Name", "Name");
                    bulkCopy.ColumnMappings.Add("Code", "Code");
                    bulkCopy.ColumnMappings.Add("Type", "Type");
                    bulkCopy.ColumnMappings.Add("Name2", "Name2");
                    bulkCopy.ColumnMappings.Add("OzelKod", "OzelKod");
                    bulkCopy.ColumnMappings.Add("Code2", "Code2");
                    bulkCopy.ColumnMappings.Add("GTIP", "GTIP");
                    bulkCopy.ColumnMappings.Add("TradeMark", "TradeMark");
         

                    bulkCopy.WriteToServer(dataTable);
                }
            }
        }
    }
}
