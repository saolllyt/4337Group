using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _4337Project
{
    public class Export_saolllyt
    {
        public static void ExportData(string connectionString, string tableName, string outputFilePath)
        {
            List<Dictionary<string, object>> data = GetDataFromTable(connectionString, tableName);

            var groupedData = data.GroupBy(row => row["Дата создания"]?.ToString());

            CreateExcel(groupedData, outputFilePath);
        }

        private static List<Dictionary<string, object>> GetDataFromTable(string connectionString, string tableName)
        {
            List<Dictionary<string, object>> data = new List<Dictionary<string, object>>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = $"SELECT * FROM {tableName}";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Dictionary<string, object> row = new Dictionary<string, object>();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                            }
                            data.Add(row);
                        }
                    }
                }
            }

            return data;
        }
        private static void CreateExcel(IEnumerable<IGrouping<string, Dictionary<string, object>>> groupedData, string outputFilePath)
        {
            FileInfo newFile = new FileInfo(outputFilePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                foreach (var group in groupedData)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(group.Key);

                    List<int> columnsToExport = new List<int> { 1, 2, 5, 6 };

                    List<string> allColumnNames = group.First().Keys.ToList();

                    int excelColumnIndex = 1;
                    foreach (int columnIndex in columnsToExport)
                    {
                        if (columnIndex <= allColumnNames.Count)
                        {
                            worksheet.Cells[1, excelColumnIndex].Value = allColumnNames[columnIndex - 1];
                            excelColumnIndex++;
                        }
                    }

                    int row = 2;
                    foreach (var record in group)
                    {
                        excelColumnIndex = 1;
                        foreach (int columnIndex in columnsToExport)
                        {
                            if (columnIndex <= allColumnNames.Count)
                            {
                                string columnName = allColumnNames[columnIndex - 1];
                                worksheet.Cells[row, excelColumnIndex].Value = record[columnName];
                                excelColumnIndex++;
                            }
                        }
                        row++;
                    }
                }

                package.Save();
            }
        }
    }
}
