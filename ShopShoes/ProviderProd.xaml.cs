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
    /// Логика взаимодействия для ProviderProd.xaml
    /// </summary>
    public partial class ProviderProd : Window
    {
        public static SqlConnection connect = new SqlConnection
            ("Data Source=laptop-1dlhhb42;Initial Catalog=ShopShoes;Integrated Security=True");
        public ProviderProd()
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
                workSheet.Name = "ProviderExp";
                System.Data.DataTable tempDt = DtIN;
                ProviderDG.ItemsSource = tempDt.DefaultView;
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

            SqlCommand command = new SqlCommand("SELECT ID_Providers, Name_Providers as 'Наименование организации' from Providers", connect);

            DataTable datatbl = new DataTable();

            datatbl.Load(command.ExecuteReader());

            ProviderDG.ItemsSource = datatbl.DefaultView;


            connect.Close();
        }

        private void ProviderDG_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ProviderDG.SelectedItem == null || ProviderDG.SelectedIndex == ProviderDG.Items.Count - 1) return;

            DataRowView row = (DataRowView)ProviderDG.SelectedItem;

            Name_provider.Text = row["Наименование организации"].ToString();
        }

        private void Insert_Click(object sender, RoutedEventArgs e)
        {
            if (Name_provider.Text == null)
            { MessageBox.Show("Все поля должны быть заполнены."); return; };
            Regex reg = new Regex(@"[0-9]");
            if (reg.IsMatch(Name_provider.Text)) { MessageBox.Show("Нельзя вводить числа."); return; }

            try
            {
                connect.Open();
                SqlCommand add = new SqlCommand("Providers_Insert", connect);

                add.CommandType = CommandType.StoredProcedure;
                add.Parameters.AddWithValue("@Name_Providers", Name_provider.Text);
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
            if (Name_provider.Text == null)
            { MessageBox.Show("Все поля должны быть заполнены."); return; };
            Regex reg = new Regex(@"[0-9]");
            if (reg.IsMatch(Name_provider.Text)) { MessageBox.Show("Нельзя вводить числа."); return; }
            DataRowView row = (DataRowView)ProviderDG.SelectedItem;

            try
            {
                connect.Open();
                SqlCommand Upd = new SqlCommand("Providers_Update", connect);
                Upd.CommandType = CommandType.StoredProcedure;
                Upd.Parameters.AddWithValue("@Name_Providers", Name_provider.Text);
                Upd.Parameters.AddWithValue("ID_Providers", (int)row["ID_Providers"]);
                Upd.ExecuteNonQuery();
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
            if (ProviderDG.SelectedItem == null) { MessageBox.Show("Выберите поле которое хотите удалить"); return; }
            if (ProviderDG.SelectedIndex == ProviderDG.Items.Count - 1) { MessageBox.Show("Вы выбрали пустое поле."); return; }
            DataRowView row = (DataRowView)ProviderDG.SelectedItem;
            try
            {

                connect.Open();
                SqlCommand Del = new SqlCommand("Providers_Delete", connect);
                Del.CommandType = CommandType.StoredProcedure;
                Del.Parameters.AddWithValue("ID_Providers", (int)row["ID_Providers"]);
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

            SqlCommand command = new SqlCommand("SELECT ID_Providers, Name_Providers as 'Наименование организации' from Providers", connect);

            DataTable datatbl = new DataTable();

            datatbl.Load(command.ExecuteReader());

            ProviderDG.ItemsSource = datatbl.DefaultView;

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

            ProviderDG.ItemsSource = readFile(openFileDialog.FileName);
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