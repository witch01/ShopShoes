using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

namespace ShopShoes
{
    /// <summary>
    /// Логика взаимодействия для Fillials.xaml
    /// </summary>
    public partial class Fillials : Window
    {
        public static SqlConnection connect = new SqlConnection
            ("Data Source=laptop-1dlhhb42;Initial Catalog=ShopShoes;Integrated Security=True");
        public Fillials()
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
                workSheet.Name = "FillialsExp";
                System.Data.DataTable tempDt = DtIN;
                FillialsDG.ItemsSource = tempDt.DefaultView;
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

            SqlCommand command = new SqlCommand("SELECT ID_Fillials, Adress as 'Адрес' from Fillials", connect);

            DataTable datatbl = new DataTable();

            datatbl.Load(command.ExecuteReader());

            FillialsDG.ItemsSource = datatbl.DefaultView;


            connect.Close();
        }

        private void FillialsDG_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FillialsDG.SelectedItem == null || FillialsDG.SelectedIndex == FillialsDG.Items.Count - 1) return;

            DataRowView row = (DataRowView)FillialsDG.SelectedItem;

            Name_type.Text = row["Адрес"].ToString();
        }

        private void Insert_Click(object sender, RoutedEventArgs e)
        {
            if (Name_type.Text == null)
            { MessageBox.Show("Все поля должны быть заполнены."); return; };
            Regex reg = new Regex(@"[0-9]");
            if (reg.IsMatch(Name_type.Text)) { MessageBox.Show("Нельзя вводить числа."); return; }

            try
            {
                connect.Open();
                SqlCommand add = new SqlCommand("Fillials_Insert", connect);

                add.CommandType = CommandType.StoredProcedure;
                add.Parameters.AddWithValue("@Adress", Name_type.Text);
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
            if (Name_type.Text == null)
            { MessageBox.Show("Все поля должны быть заполнены."); return; };
            Regex reg = new Regex(@"[0-9]");
            if (reg.IsMatch(Name_type.Text)) { MessageBox.Show("Нельзя вводить числа."); return; }
            DataRowView row = (DataRowView)FillialsDG.SelectedItem;

            try
            {
                connect.Open();
                SqlCommand Upd = new SqlCommand("Fillials_Update", connect);
                Upd.CommandType = CommandType.StoredProcedure;
                Upd.Parameters.AddWithValue("@Adress", Name_type.Text);
                Upd.Parameters.AddWithValue("ID_Fillials", (int)row["ID_Fillials"]);
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
            if (FillialsDG.SelectedItem == null) { MessageBox.Show("Выберите поле которое хотите удалить"); return; }
            if (FillialsDG.SelectedIndex == FillialsDG.Items.Count - 1) { MessageBox.Show("Вы выбрали пустое поле."); return; }
            DataRowView row = (DataRowView)FillialsDG.SelectedItem;
            try
            {

                connect.Open();
                SqlCommand Del = new SqlCommand("Fillials_Delete", connect);
                Del.CommandType = CommandType.StoredProcedure;
                Del.Parameters.AddWithValue("ID_Fillials", (int)row["ID_Fillials"]);
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

            SqlCommand command = new SqlCommand("SELECT ID_Fillials, Adress as 'Адрес' from Fillials", connect);

            DataTable datatbl = new DataTable();

            datatbl.Load(command.ExecuteReader());

            FillialsDG.ItemsSource = datatbl.DefaultView;
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

            FillialsDG.ItemsSource = readFile(openFileDialog.FileName);
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
            Menu menu = new Menu();
            menu.Show();
            this.Close();
        }
    }
    }

