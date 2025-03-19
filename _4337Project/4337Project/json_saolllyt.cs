using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Newtonsoft.Json;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace _4337Project
{
    internal class json_saolllyt
    {
       private string connectionString = "Server=SAOLLLYT;Database=data_saolllyt;Integrated Security=True;";
        public void CreateOrdersTable()
        {
            string query = @"
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Orders' AND xtype='U')
            CREATE TABLE Orders (
                Id INT PRIMARY KEY,
                CodeOrder NVARCHAR(50),
                CreateDate NVARCHAR(10),
                CreateTime NVARCHAR(5),
                CodeClient NVARCHAR(50),
                Services NVARCHAR(MAX),
                Status NVARCHAR(50),
                ClosedDate NVARCHAR(10),
                ProkatTime NVARCHAR(50)
            );";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(query, connection);
                command.ExecuteNonQuery();
            }
        }

        public List<Order> LoadOrdersFromJson(string filePath)
        {
            List<Order> orders = new List<Order>();

            if (File.Exists(filePath))
            {
                string json = File.ReadAllText(filePath);
                orders = JsonConvert.DeserializeObject<List<Order>>(json);
            }
            else
            {
                MessageBox.Show("Файл не найден");
            }

            return orders;
        }

        public void SaveOrdersToDatabase(List<Order> orders)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string deleteQuery = "DELETE FROM Orders";
                SqlCommand deleteCmd = new SqlCommand(deleteQuery, connection);
                deleteCmd.ExecuteNonQuery();

                foreach (var order in orders)
                {
                    string query = @"
            INSERT INTO Orders (Id, CodeOrder, CreateDate, CreateTime, CodeClient, Services, Status, ClosedDate, ProkatTime)
            VALUES (@Id, @CodeOrder, @CreateDate, @CreateTime, @CodeClient, @Services, @Status, @ClosedDate, @ProkatTime)";

                    SqlCommand command = new SqlCommand(query, connection);
                    command.Parameters.AddWithValue("@Id", order.Id);
                    command.Parameters.AddWithValue("@CodeOrder", order.CodeOrder);
                    command.Parameters.AddWithValue("@CreateDate", order.CreateDate);
                    command.Parameters.AddWithValue("@CreateTime", order.CreateTime);
                    command.Parameters.AddWithValue("@CodeClient", order.CodeClient);
                    command.Parameters.AddWithValue("@Services", order.Services);
                    command.Parameters.AddWithValue("@Status", order.Status);
                    command.Parameters.AddWithValue("@ClosedDate", order.ClosedDate ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@ProkatTime", order.ProkatTime);

                    command.ExecuteNonQuery();
                }
            }
        }


        public Dictionary<string, List<Order>> GroupOrdersByStatus(List<Order> orders)
        {
            return orders
                .GroupBy(order => order.Status)
                .ToDictionary(group => group.Key, group => group.ToList());
        }

        public void ExportToWord(Dictionary<string, List<Order>> groupedOrders, string filePath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = new Body();

                foreach (var group in groupedOrders)
                {
                    Paragraph statusParagraph = new Paragraph();
                    Run statusRun = new Run();
                    Text statusText = new Text($"Статус: {group.Key}");
                    statusRun.Append(statusText);
                    statusParagraph.Append(statusRun);
                    body.Append(statusParagraph);

                    Table table = new Table();

                    TableRow headerRow = new TableRow();
                    string[] headers = { "Id", "Код заказа", "Дата создания", "Код клиента", "Услуги" };
                    foreach (var header in headers)
                    {
                        TableCell cell = new TableCell(new Paragraph(new Run(new Text(header))));
                        headerRow.Append(cell);
                    }
                    table.Append(headerRow);

                    foreach (var order in group.Value)
                    {
                        TableRow row = new TableRow();

                        var cells = new[]
                        {
        new TableCell(new Paragraph(new Run(new Text(order.Id.ToString())))),
        new TableCell(new Paragraph(new Run(new Text(order.CodeOrder)))),
        new TableCell(new Paragraph(new Run(new Text(order.CreateDate)))),
        new TableCell(new Paragraph(new Run(new Text(order.CodeClient)))),
        new TableCell(new Paragraph(new Run(new Text(order.Services))))
    };

                        foreach (var cell in cells)
                        {
                            row.Append(cell);
                        }

                        table.Append(row);
                    }


                    body.Append(table);
                    body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                }

                mainPart.Document.Append(body);
                mainPart.Document.Save();
            }
        }

        public List<Order> LoadOrdersFromDatabase()
        {
            List<Order> orders = new List<Order>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT * FROM Orders";
                SqlCommand command = new SqlCommand(query, connection);

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        orders.Add(new Order
                        {
                            Id = reader.GetInt32(0),
                            CodeOrder = reader.GetString(1),
                            CreateDate = reader.GetString(2),
                            CreateTime = reader.GetString(3),
                            CodeClient = reader.GetString(4),
                            Services = reader.GetString(5),
                            Status = reader.GetString(6),
                            ClosedDate = reader.IsDBNull(7) ? null : reader.GetString(7),
                            ProkatTime = reader.GetString(8)
                        });
                    }
                }
            }

            return orders;
        }
    }
}
