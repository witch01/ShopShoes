using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
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
    /// Логика взаимодействия для AviabillityProd.xaml
    /// </summary>
    public partial class AviabillityProd : Window
    {
        public static SqlConnection connect = new SqlConnection
          ("Data Source=laptop-1dlhhb42;Initial Catalog=ShopShoes;Integrated Security=True");
        public AviabillityProd()
        {
            InitializeComponent();
        }

        Microsoft.Office.Interop.Excel.Application excel;
        Microsoft.Office.Interop.Excel.Workbook workBook;
        Microsoft.Office.Interop.Excel.Worksheet workSheet;
        Microsoft.Office.Interop.Excel.Range cellRange;
        IExcelDataReader edr;
        private void GenerateExcel(DataTable DtIN)
        {
            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.DisplayAlerts = false;
                excel.Visible = false;
                workBook = excel.Workbooks.Add(Type.Missing);
                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                workSheet.Name = "AviabillityExp";
                System.Data.DataTable tempDt = DtIN;
                AviabillityDG.ItemsSource = tempDt.DefaultView;
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

            SqlCommand command = new SqlCommand("SELECT ID_Availability, Name_Product as 'Товар', [Adress] AS 'Филиал' from [dbo].[Availability] join [dbo].[Product] on [ID_Product]=[Product_ID] join [dbo].[Fillials] on [ID_Fillials]=[Fillials_ID]", connect);

            DataTable datatbl = new DataTable();

            datatbl.Load(command.ExecuteReader());

            AviabillityDG.ItemsSource = datatbl.DefaultView;

            SqlCommand command1 = new SqlCommand("SELECT ID_Product, Name_Product from [dbo].[Product]", connect);
            DataTable datatbl1 = new DataTable();
            datatbl1.Load(command1.ExecuteReader());
            Product.ItemsSource = datatbl1.DefaultView;
            Product.DisplayMemberPath = "Name_Product";
            Product.SelectedValuePath = "ID_Product";

            SqlCommand command2 = new SqlCommand("SELECT ID_Fillials, Adress from [dbo].[Fillials]", connect);
            DataTable datatbl2 = new DataTable();
            datatbl2.Load(command2.ExecuteReader());
            Fillials.ItemsSource = datatbl2.DefaultView;
            Fillials.DisplayMemberPath = "Adress";
            Fillials.SelectedValuePath = "ID_Fillials";

            connect.Close();
        }

        private void AviabillityDG_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (AviabillityDG.SelectedItem == null || AviabillityDG.SelectedIndex == AviabillityDG.Items.Count - 1) return;

            DataRowView row = (DataRowView)AviabillityDG.SelectedItem;

            Product.Text = row["Товар"].ToString();
            Fillials.Text = row["Филиал"].ToString();
        }

        private void Insert_Click(object sender, RoutedEventArgs e)
        {
            if (Product.Text == null || Fillials.Text == null)
            { MessageBox.Show("Все поля должны быть заполнены."); return; };
            try
            {
                connect.Open();
                SqlCommand add = new SqlCommand("Availability_Insert", connect);

                add.CommandType = CommandType.StoredProcedure;
                add.Parameters.AddWithValue("@Product_ID", Product.SelectedValue);
                add.Parameters.AddWithValue("@Fillials_ID", Fillials.SelectedValue);
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
            if (Product.Text == null || Fillials.Text == null)
            { MessageBox.Show("Все поля должны быть заполнены."); return; };
            DataRowView row = (DataRowView)AviabillityDG.SelectedItem;
            try
            {
                connect.Open();
                SqlCommand upd = new SqlCommand("Availability_Update", connect);

                upd.CommandType = CommandType.StoredProcedure;
                upd.Parameters.AddWithValue("@Product_ID", Product.SelectedValue);
                upd.Parameters.AddWithValue("@Fillials_ID", Fillials.SelectedValue);
                upd.Parameters.AddWithValue("ID_Availability", (int)row["ID_Availability"]);
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
            if (AviabillityDG.SelectedItem == null) { MessageBox.Show("Выберите поле которое хотите удалить"); return; }
            if (AviabillityDG.SelectedIndex == AviabillityDG.Items.Count - 1) { MessageBox.Show("Вы выбрали пустое поле."); return; }
            DataRowView row = (DataRowView)AviabillityDG.SelectedItem;
            try
            {

                connect.Open();
                SqlCommand Del = new SqlCommand("Availability_Delete", connect);
                Del.CommandType = CommandType.StoredProcedure;
                Del.Parameters.AddWithValue("ID_Availability", (int)row["ID_Availability"]);
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

            SqlCommand command = new SqlCommand("SELECT ID_Availability, Name_Product as 'Товар', [Adress] AS 'Филиал' from [dbo].[Availability] join [dbo].[Product] on [ID_Product]=[Product_ID] join [dbo].[Fillials] on [ID_Fillials]=[Fillials_ID]", connect);

            DataTable datatbl = new DataTable();

            datatbl.Load(command.ExecuteReader());

            AviabillityDG.ItemsSource = datatbl.DefaultView;
            GenerateExcel(datatbl);
            connect.Close();
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EXCEL Files (*.xlsx)|*.xlsx|EXCEL Files 2003 (*.xls)|*.xls|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() != true)
                return;

            AviabillityDG.ItemsSource = readFile(openFileDialog.FileName);
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
            MenuProd menuprod = new MenuProd();
            menuprod.Show();
            this.Close();
        }
    }
}

