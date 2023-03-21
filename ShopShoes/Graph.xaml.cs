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
using System.Data;
using LiveCharts;
using LiveCharts.Wpf;
using System.Data.SqlClient;

namespace ShopShoes
{
    /// <summary>
    /// Логика взаимодействия для Graph.xaml
    /// </summary>
    public partial class Graph : Window
    {
        private SqlConnection sqlConnection = null;
        private SqlDataAdapter dataAdapter = null;
        private DataSet dataSet = null;
        private DataTable table = null;
        public Graph()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            SqlConnection connect = new SqlConnection("Data Source=laptop-1dlhhb42;Initial Catalog=ShopShoes;Integrated Security=True");
            connect.Open();
            dataAdapter = new SqlDataAdapter("SELECT * FROM Product", connect); // исправляем ошибку с подключением к БД
            dataSet = new DataSet();
            dataAdapter.Fill(dataSet, "Product");
            table = dataSet.Tables["Product"];
            graf.LegendLocation = LegendLocation.Bottom;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Window s = new Product();
            this.Hide();
            s.Show();
        }


        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            SeriesCollection series = new SeriesCollection();
            ChartValues<string> positionValues = new ChartValues<string>();
            ChartValues<double> salaries = new ChartValues<double>(); // изменяем тип на ChartValues<double>
            foreach (DataRow row in table.Rows)
            {
                positionValues.Add(Convert.ToString(row["Name_Product"]));
                salaries.Add(Double.Parse(row["Cent"].ToString().Replace('.', ',')));
            }
            graf.AxisX.Clear();
            graf.AxisX.Add(new Axis()
            {
                Title = "Товары",
                Labels = positionValues
            });

            ColumnSeries column = new ColumnSeries();
            column.Title = "Стоимость";
            column.Values = salaries;

            series.Add(column);
            graf.Series = series;
        }
    }
}