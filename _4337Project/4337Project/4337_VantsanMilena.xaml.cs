using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace _4337Project
{
    /// <summary>
    /// Логика взаимодействия для _4337_VantsanMilena.xaml
    /// </summary>
    public partial class _4337_VantsanMilena : Window
    {
        string connectionString = "Server=SAOLLLYT;Database=data_saolllyt;Integrated Security=True;";

        public _4337_VantsanMilena()
        {
            InitializeComponent();
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
                string tableName = "Table_" + fileName;
                try
                {
                    Import_saolllyt.ImportData(filePath, connectionString, tableName);
                    MessageBox.Show("Данные успешно импортированы!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при импорте данных: {ex.Message}");
                }
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            string tableName = "Table_2";

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.FileName = "ExportBD.xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {
                string outputFilePath = saveFileDialog.FileName;

                try
                {
                    Export_saolllyt.ExportData(connectionString, tableName, outputFilePath);
                    MessageBox.Show("Данные успешно экспортированы по дате создания заказа");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при экспорте данных: {ex.Message}");
                }
            }

        }

        private void ImportJSON_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;

                json_saolllyt orderService = new json_saolllyt();
                orderService.CreateOrdersTable();

                List<Order> orders = orderService.LoadOrdersFromJson(filePath);

                if (orders != null && orders.Count > 0)
                {
                    orderService.SaveOrdersToDatabase(orders);
                    MessageBox.Show("Данные успешно импортированы");
                }
                else
                {
                    MessageBox.Show("Не удалось загрузить данные");
                }
            }
            else
            {
                MessageBox.Show("Файл не выбран");
            }
        }

        private void ExportWORD_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word files (*.docx)|*.docx|All files (*.*)|*.*";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.FileName = "ExportBD.docx";

            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;

                json_saolllyt orderService = new json_saolllyt();
                List<Order> orders = orderService.LoadOrdersFromDatabase();

                if (orders != null && orders.Count > 0)
                {
                    Dictionary<string, List<Order>> groupedOrders = orderService.GroupOrdersByStatus(orders);
                    orderService.ExportToWord(groupedOrders, filePath);

                    MessageBox.Show("Данные успешно экспортированы в Word");
                }
                else
                {
                    MessageBox.Show("Нет данных для экспорта");
                }
            }
        }
    }
}
