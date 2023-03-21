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
    /// Логика взаимодействия для MenuProd.xaml
    /// </summary>
    public partial class MenuProd : Window
    {
        public MenuProd()
        {
            InitializeComponent();
        }
        private void Aviabillity_Click(object sender, RoutedEventArgs e)
        {
            AviabillityProd aviabillity = new AviabillityProd();
            aviabillity.Show();
            this.Close();
        }
        private void Order_Click(object sender, RoutedEventArgs e)
        {
            OrderProd order = new OrderProd(); ;
            order.Show();
            this.Close();
        }

        private void Product_Click(object sender, RoutedEventArgs e)
        {
            ProductProd product = new ProductProd();
            product.Show();
            this.Close();
        }

        private void Provider_Click(object sender, RoutedEventArgs e)
        {
            ProviderProd provider = new ProviderProd();
            provider.Show();
            this.Close();
        }

        private void TypeProduct_Click(object sender, RoutedEventArgs e)
        {
            TypeproductProd type_Product = new TypeproductProd();
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
