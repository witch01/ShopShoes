using System;
using System.Collections.Generic;
using System.Data.SqlClient;
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
    /// Логика взаимодействия для Login.xaml
    /// </summary>
    public partial class Login : Window
    {
        public static SqlConnection connect = new SqlConnection
           ("Data Source=laptop-1dlhhb42;Initial Catalog=ShopShoes;Integrated Security=True");
        public Login()
        {
            InitializeComponent();
        }
        private void Vxod_Click(object sender, RoutedEventArgs e)
        {
            connect.Open();
            SqlCommand c = new SqlCommand($@"SELECT [Phone],[Passwd],Post_ID FROM [Employee] where [Phone]='{Phone.Text}' and Passwd='{Password.Password}'", connect);
            SqlDataReader reader = c.ExecuteReader();
            if (reader.Read())
            {
                string role = Convert.ToString(reader["Post_ID"].ToString());
                if (reader.HasRows) 
                {

                    switch (role)
                    {
                        case "2":
                            {
                                Window u = new MenuProd();
                                this.Hide();
                                u.Show();
                                break;
                            }
                        case "3":
                            {

                                Window u1 = new Menu();
                                this.Hide();
                                u1.Show();
                                break;
                            }
                      

                    }
                }
                


            }
            else { MessageBox.Show("Пользователь не найден"); }
            connect.Close();
        }
    }
}
