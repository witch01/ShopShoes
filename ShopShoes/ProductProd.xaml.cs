using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
    /// Логика взаимодействия для ProductProd.xaml
    /// </summary>
    public partial class ProductProd : Window
    {
        public static SqlConnection connect = new SqlConnection
            ("Data Source=laptop-1dlhhb42;Initial Catalog=ShopShoes;Integrated Security=True");
        public ProductProd()
        {
            InitializeComponent();
        }
        Microsoft.Office.Interop.Excel.Application excel;
        Microsoft.Office.Interop.Excel.Workbook workBook;
        Microsoft.Office.Interop.Excel.Worksheet workSheet;
        Microsoft.Office.Interop.Excel.Range cellRange;
        private void GenerateExcel(DataTable DtIN)
        {
            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.DisplayAlerts = false;
                excel.Visible = false;
                workBook = excel.Workbooks.Add(Type.Missing);
                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                workSheet.Name = "ProductExp";
                System.Data.DataTable tempDt = DtIN;
                ProductDG.ItemsSource = tempDt.DefaultView;
                workSheet.Cells.Font.Size = 11;
                int rowcount = 1;
                for (int i = 1; i <= tempDt.Columns.Count; i++) //taking care of Headers.
                {
                    workSheet.Cells[1, i] = tempDt.Columns[i - 1].ColumnName;
                }
                foreach (System.Data.DataRow row in tempDt.Rows) //taking care of each Row
                {
                    rowcount += 1;
                    for (int i = 0; i < tempDt.Columns.Count; i++) //taking care of each column
                    {
                        workSheet.Cells[rowcount, i + 1] = row[i].ToString();
                    }
                }
                cellRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowcount, tempDt.Columns.Count]];
                cellRange.EntireColumn.AutoFit();
                excel.Visible = true;
                excel.UserControl = true;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            connect.Open();

            SqlCommand command = new SqlCommand("SELECT ID_Product, Name_Product as 'Наименование', " +
                "[Cent] AS 'Цена', Quantity AS 'Остаток', Datepr as 'Дата производства', Name_Type_product as 'Категория', Name_Providers as 'Поставщик' from Product " +
                "join [dbo].[Type_product] on [ID_Type_product]=[Type_product_ID] join [dbo].[Providers] on [ID_Providers]=[Providers_ID]", connect);

            DataTable datatbl = new DataTable();

            datatbl.Load(command.ExecuteReader());

            ProductDG.ItemsSource = datatbl.DefaultView;

            SqlCommand command1 = new SqlCommand("SELECT ID_Type_product, Name_Type_product from [dbo].[Type_product]", connect);
            DataTable datatbl1 = new DataTable();
            datatbl1.Load(command1.ExecuteReader());
            Type_product.ItemsSource = datatbl1.DefaultView;
            Type_product.DisplayMemberPath = "Name_Type_product";
            Type_product.SelectedValuePath = "ID_Type_product";

            SqlCommand command2 = new SqlCommand("SELECT ID_Providers, Name_Providers from [dbo].[Providers]", connect);
            DataTable datatbl2 = new DataTable();
            datatbl2.Load(command2.ExecuteReader());
            Provider.ItemsSource = datatbl2.DefaultView;
            Provider.DisplayMemberPath = "Name_Providers";
            Provider.SelectedValuePath = "ID_Providers";

            connect.Close();
        }

        private void ProductDG_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ProductDG.SelectedItem == null || ProductDG.SelectedIndex == ProductDG.Items.Count - 1) return;

            DataRowView row = (DataRowView)ProductDG.SelectedItem;

            Name.Text = row["Наименование"].ToString();
            Cent.Text = row["Цена"].ToString();
            Quantity.Text = row["Остаток"].ToString();
            DR.Text = row["Дата производства"].ToString();
            Type_product.Text = row["Категория"].ToString();
            Provider.Text = row["Поставщик"].ToString();
        }

        private void Insert_Click(object sender, RoutedEventArgs e)
        {
            if (Name.Text == null || Cent.Text == null || Quantity.Text == null || DR.Text == null || Type_product.Text == null || Provider.Text == null)
            { MessageBox.Show("Все поля должны быть заполнены."); return; };
            Regex reg = new Regex(@"[0-9]");
            if (reg.IsMatch(Name.Text)) { MessageBox.Show("Нельзя вводить числа."); return; }
            Regex reg1 = new Regex(@"[A-Z]");
            Regex reg2 = new Regex(@"[a-z]");
            Regex reg3 = new Regex(@"[А-Я]");
            Regex reg4 = new Regex(@"[а-я]");
            if (reg1.IsMatch(Cent.Text) || reg2.IsMatch(Cent.Text) || reg3.IsMatch(Cent.Text) || reg4.IsMatch(Cent.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            if (reg1.IsMatch(Quantity.Text) || reg2.IsMatch(Quantity.Text) || reg3.IsMatch(Quantity.Text) || reg4.IsMatch(Quantity.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            try
            {
                connect.Open();
                SqlCommand add = new SqlCommand("Product_Insert", connect);

                add.CommandType = CommandType.StoredProcedure;
                add.Parameters.AddWithValue("@Name_Product", Name.Text);
                add.Parameters.AddWithValue("@Cent", Convert.ToDecimal(Cent.Text));
                add.Parameters.AddWithValue("@Quantity", Convert.ToInt32(Quantity.Text));
                add.Parameters.AddWithValue("@Datepr", DR.SelectedDate);
                add.Parameters.AddWithValue("@Type_product_ID", Type_product.SelectedValue);
                add.Parameters.AddWithValue("@Providers_ID", Provider.SelectedValue);
                add.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }

            finally
            {
                connect.Close();
                Window_Loaded(sender, e);
            }
        }

        private void Update_Click(object sender, RoutedEventArgs e)
        {
            if (Name.Text == null || Cent.Text == null || Quantity.Text == null || DR.Text == null || Type_product.Text == null || Provider.Text == null)
            { MessageBox.Show("Все поля должны быть заполнены."); return; };
            Regex reg = new Regex(@"[0-9]");
            if (reg.IsMatch(Name.Text)) { MessageBox.Show("Нельзя вводить числа."); return; }
            Regex reg1 = new Regex(@"[A-Z]");
            Regex reg2 = new Regex(@"[a-z]");
            Regex reg3 = new Regex(@"[А-Я]");
            Regex reg4 = new Regex(@"[а-я]");
            if (reg1.IsMatch(Cent.Text) || reg2.IsMatch(Cent.Text) || reg3.IsMatch(Cent.Text) || reg4.IsMatch(Cent.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            if (reg1.IsMatch(Quantity.Text) || reg2.IsMatch(Quantity.Text) || reg3.IsMatch(Quantity.Text) || reg4.IsMatch(Quantity.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            DataRowView row = (DataRowView)ProductDG.SelectedItem;
            try
            {
                connect.Open();
                SqlCommand upd = new SqlCommand("Product_Update", connect);

                upd.CommandType = CommandType.StoredProcedure;
                upd.Parameters.AddWithValue("@Name_Product", Name.Text);
                upd.Parameters.AddWithValue("@Cent", Convert.ToDecimal(Cent.Text));
                upd.Parameters.AddWithValue("@Quantity", Convert.ToInt32(Quantity.Text));
                upd.Parameters.AddWithValue("@Datepr", DR.SelectedDate);
                upd.Parameters.AddWithValue("@Type_product_ID", Type_product.SelectedValue);
                upd.Parameters.AddWithValue("@Providers_ID", Provider.SelectedValue);
                upd.Parameters.AddWithValue("ID_Product", (int)row["ID_Product"]);
                upd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }

            finally
            {
                connect.Close();
                Window_Loaded(sender, e);
            }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (ProductDG.SelectedItem == null) { MessageBox.Show("Выберите поле которое хотите удалить"); return; }
            if (ProductDG.SelectedIndex == ProductDG.Items.Count - 1) { MessageBox.Show("Вы выбрали пустое поле."); return; }
            DataRowView row = (DataRowView)ProductDG.SelectedItem;
            try
            {

                connect.Open();
                SqlCommand Del = new SqlCommand("Product_Delete", connect);
                Del.CommandType = CommandType.StoredProcedure;
                Del.Parameters.AddWithValue("ID_Product", (int)row["ID_Product"]);
                Del.ExecuteNonQuery();
            }

            catch (SqlException ex)

            {
                MessageBox.Show(ex.Message);
            }

            finally
            {
                connect.Close();
                Window_Loaded(sender, e);
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            connect.Open();

            SqlCommand command = new SqlCommand("SELECT ID_Product, Name_Product as 'Наименование', " +
                "[Cent] AS 'Цена', Quantity AS 'Остаток', Datepr as 'Дата производства', Name_Type_product as 'Категория', Name_Providers as 'Поставщик' from Product " +
                "join [dbo].[Type_product] on [ID_Type_product]=[Type_product_ID] join [dbo].[Providers] on [ID_Providers]=[Providers_ID]", connect);

            DataTable datatbl = new DataTable();

            datatbl.Load(command.ExecuteReader());

            ProductDG.ItemsSource = datatbl.DefaultView;
            GenerateExcel(datatbl);
            connect.Close();
        }

        IExcelDataReader edr;
        private void Import_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EXCEL Files (*.xlsx)|*.xlsx|EXCEL Files 2003 (*.xls)|*.xls|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() != true)
                return;

            ProductDG.ItemsSource = readFile(openFileDialog.FileName);
        }
        private DataView readFile(string fileNames)
        {

            var extension = fileNames.Substring(fileNames.LastIndexOf('.'));
            // Создаем поток для чтения.
            FileStream stream = File.Open(fileNames, FileMode.Open, FileAccess.Read);
            // В зависимости от расширения файла Excel, создаем тот или иной читатель.
            // Читатель для файлов с расширением *.xlsx.
            if (extension == ".xlsx")
                edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
            // Читатель для файлов с расширением *.xls.
            else if (extension == ".xls")
                edr = ExcelReaderFactory.CreateBinaryReader(stream);

            //// reader.IsFirstRowAsColumnNames
            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };
            // Читаем, получаем DataView и работаем с ним как обычно.
            DataSet dataSet = edr.AsDataSet(conf);
            DataView dtView = dataSet.Tables[0].AsDataView();

            // После завершения чтения освобождаем ресурсы.
            edr.Close();
            return dtView;
        }

        private void Menu_Click(object sender, RoutedEventArgs e)
        {
            MenuProd menu = new MenuProd();
            menu.Show();
            this.Close();
        }
    }
}
