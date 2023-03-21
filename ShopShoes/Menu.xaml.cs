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

namespace ShopShoes
{
    /// <summary>
    /// Логика взаимодействия для Menu.xaml
    /// </summary>
    public partial class Menu : Window
    {
        public Menu()
        {
            InitializeComponent();
        }

        private void Aviabillity_Click(object sender, RoutedEventArgs e)
        {
            Aviabillity aviabillity = new Aviabillity();
            aviabillity.Show();
            this.Close();
        }

        private void Employee_Click(object sender, RoutedEventArgs e)
        {
            Employee employee = new Employee();
            employee.Show();
            this.Close();
        }

        private void Fillials_Click(object sender, RoutedEventArgs e)
        {
            Fillials fillials = new Fillials();;
            fillials.Show();
            this.Close();
        }

        private void Post_Click(object sender, RoutedEventArgs e)
        {
            MainWindow post = new MainWindow();
            post.Show();
            this.Close();
        }

        private void Order_Click(object sender, RoutedEventArgs e)
        {
            Order order=new Order();;
            order.Show();
            this.Close();
        }

        private void Product_Click(object sender, RoutedEventArgs e)
        {
            Product product = new Product();
            product.Show();
            this.Close();
        }

        private void Provider_Click(object sender, RoutedEventArgs e)
        {
            Provider provider = new Provider();
            provider.Show();
            this.Close();
        }

        private void TypeProduct_Click(object sender, RoutedEventArgs e)
        {
            Type_product type_Product = new Type_product();
            type_Product.Show();
            this.Close();
        }

        private void Avto_Click(object sender, RoutedEventArgs e)
        {
            Login login = new Login();
            login.Show();
            this.Close();
        }
    }
}
