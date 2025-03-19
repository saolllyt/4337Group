using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.IO;

namespace _4337Project
{
    public class Import_saolllyt
    {
        public static void ImportData(string filePath, string connectionString, string tableName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                List<string> columnNames = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    columnNames.Add(worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}");
                }

                CreateTable(connectionString, tableName, columnNames);

                SaveDataToTable(connectionString, tableName, worksheet, columnNames);
            }
        }

        public static void CreateTable(string connectionString, string tableName, List<string> columnNames)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string createTableQuery = $"IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = '{tableName}') BEGIN CREATE TABLE {tableName} (";

                for (int i = 0; i < columnNames.Count; i++)
                {
                    createTableQuery += $"[{columnNames[i]}] NVARCHAR(MAX)";
                    if (i < columnNames.Count - 1)
                    {
                        createTableQuery += ", ";
                    }
                }

                createTableQuery += ") END";

                using (SqlCommand command = new SqlCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public static void SaveDataToTable(string connectionString, string tableName, ExcelWorksheet worksheet, List<string> columnNames)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                for (int row = 2; row <= rowCount; row++)
                {
                    bool isEmptyRow = true;
                    for (int col = 1; col <= colCount; col++)
                    {
                        if (worksheet.Cells[row, col].Value != null && !string.IsNullOrWhiteSpace(worksheet.Cells[row, col].Value.ToString()))
                        {
                            isEmptyRow = false;
                            break;
                        }
                    }

                    if (!isEmptyRow)
                    {
                        string insertQuery = $"INSERT INTO {tableName} (";
                        for (int i = 0; i < columnNames.Count; i++)
                        {
                            insertQuery += $"[{columnNames[i]}]";
                            if (i < columnNames.Count - 1)
                            {
                                insertQuery += ",";
                            }
                        }
                        insertQuery += ") VALUES (";
                        for (int i = 0; i < columnNames.Count; i++)
                        {
                            insertQuery += $"@p{i}";
                            if (i < columnNames.Count - 1)
                            {
                                insertQuery += ",";
                            }
                        }
                        insertQuery += ")";

                        using (SqlCommand command = new SqlCommand(insertQuery, connection))
                        {
                            for (int i = 0; i < columnNames.Count; i++)
                            {
                                string cellValue = worksheet.Cells[row, i + 1].Value?.ToString();
                                command.Parameters.AddWithValue($"@p{i}", (object)cellValue ?? DBNull.Value);
                            }

                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
        }
    }
}
