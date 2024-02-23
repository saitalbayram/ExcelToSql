using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace ExcelToSql
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.Write("Sunucu Adı/Ip Adresi: ");
            string serverName = Console.ReadLine();
            Console.Write("Veritabanı Adı: ");
            string databaseName = Console.ReadLine();
            Console.Write("Kullanıcı Adı: ");
            string userName = Console.ReadLine();
            Console.Write("Şifre: ");
            string password = Console.ReadLine();


            string connectionString;

            if(string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(password))
            {
                connectionString = "Data Source=" + serverName + ";Initial Catalog=" + databaseName + ";Integrated Security=True;";
            }
            else
            {
                connectionString = "Data Source=" + serverName + ";Initial Catalog="+ databaseName+ "; User Id=" + userName + ";Password=" + password + ";";
            }

            if (TestConnection(connectionString))
            {
                Console.WriteLine("Sunucuyla bağlantı başarıyla kuruldu");
                Console.Write("\nExcel dosyasının dosya yolu: ");
                string excelFilePath = Console.ReadLine();

                
                // Excel'den verileri oku
                System.Data.DataTable dataTable = ReadExcelData(excelFilePath);

                // SQL Server'a verileri aktar
                WriteToSqlServer(dataTable, connectionString);

                Console.WriteLine("Veri aktarımı tamamlandı.");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine("Veritabanı bağlantısı başarısız! Lütfen bilgileri kontrol edip tekrar deneyin.");
                Console.ReadLine();
            }

          
        }

        static bool TestConnection(string connectionString)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    return true;
                }

                catch (Exception error)
                {
                    Console.WriteLine(error.Message);
                    return false;
                }

            }
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
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                    {
                        Console.Write("\nSql tablosunun adı: ");
                        string sqlTableName = Console.ReadLine();
                        bulkCopy.DestinationTableName = sqlTableName;

                        // SQL tablosunun sütun isimlerini alın
                        System.Data.DataTable schemaTable = connection.GetSchema("Columns", new[] { null, null, sqlTableName });

                        Console.WriteLine("\nKopyalanacak sütunlar:");
                        foreach (DataColumn column in dataTable.Columns)
                        {
                            // Excel sütun başlığı ile SQL tablosundaki sütun isimlerini eşleştir
                            foreach (DataRow row in schemaTable.Rows)
                            {
                                string columnName = row["COLUMN_NAME"].ToString();
                                if (string.Equals(column.ColumnName, columnName, StringComparison.OrdinalIgnoreCase))
                                {
                                    Console.WriteLine($"Excel Sütun Adı: {column.ColumnName} - SQL Sütun Adı: {columnName}");
                                    bulkCopy.ColumnMappings.Add(column.ColumnName, columnName);
                                    break;
                                }
                            }
                        }

                        bulkCopy.WriteToServer(dataTable);
                    }
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                return;
            }
        }

    }
}
