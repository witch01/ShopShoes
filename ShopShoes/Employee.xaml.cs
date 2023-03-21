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
    /// Логика взаимодействия для Employee.xaml
    /// </summary>
    public partial class Employee : Window
    {

        public static SqlConnection connect = new SqlConnection
           ("Data Source=laptop-1dlhhb42;Initial Catalog=ShopShoes;Integrated Security=True");
        public Employee()
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
                workSheet.Name = "EmployeeExp";
                System.Data.DataTable tempDt = DtIN;
                EmployeeDG.ItemsSource = tempDt.DefaultView;
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
        private void EmployeeDG_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (EmployeeDG.SelectedItem == null || EmployeeDG.SelectedIndex == EmployeeDG.Items.Count - 1) return;

            DataRowView row = (DataRowView)EmployeeDG.SelectedItem;

            Surname.Text = row["Фамилия"].ToString();
            Name.Text = row["Имя"].ToString();
            Middle_name.Text = row["Отчество"].ToString();
            Phone.Text = row["Номер телефона"].ToString();
            INN.Text = row["ИНН"].ToString();
            SNILS.Text = row["СНИЛС"].ToString();
            Passport_number_Employee.Text = row["Номер паспорта"].ToString();
            Passport_series_Employee.Text = row["Серия паспорта"].ToString();
            Passwd.Text = row["Пароль"].ToString();
            Post.Text = row["Должность"].ToString();
            DR.Text = row["Дата рождения"].ToString();
        }

        private void Insert_Click(object sender, RoutedEventArgs e)
        {
            if (Surname.Text == null || Name.Text == null || Middle_name.Text == null || Phone.Text == null || INN.Text == null ||
                SNILS.Text == null|| Passport_number_Employee.Text==null|| Passport_series_Employee.Text==null|| Passwd.Text==null|| Post.Text==null||
                DR.Text==null)
            { MessageBox.Show("Все поля должны быть заполнены."); return; };
            Regex reg = new Regex(@"[0-9]");
            if (reg.IsMatch(Name.Text)||reg.IsMatch(Surname.Text)||reg.IsMatch(Middle_name.Text)) 
            { MessageBox.Show("Нельзя вводить числа."); return; }
            Regex reg1 = new Regex(@"[A-Z]");
            Regex reg2 = new Regex(@"[a-z]");
            Regex reg3 = new Regex(@"[А-Я]");
            Regex reg4 = new Regex(@"[а-я]");
            if (reg1.IsMatch(Phone.Text) || reg2.IsMatch(Phone.Text) || reg3.IsMatch(Phone.Text) || reg4.IsMatch(Phone.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            if (reg1.IsMatch(INN.Text) || reg2.IsMatch(INN.Text) || reg3.IsMatch(INN.Text) || reg4.IsMatch(INN.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            if (reg1.IsMatch(SNILS.Text) || reg2.IsMatch(SNILS.Text) || reg3.IsMatch(SNILS.Text) || reg4.IsMatch(SNILS.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            if (reg1.IsMatch(Passport_number_Employee.Text) || reg2.IsMatch(Passport_number_Employee.Text) || reg3.IsMatch(Passport_number_Employee.Text) || reg4.IsMatch(Passport_number_Employee.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            if (reg1.IsMatch(Passport_series_Employee.Text) || reg2.IsMatch(Passport_series_Employee.Text) || reg3.IsMatch(Passport_series_Employee.Text) || reg4.IsMatch(Passport_series_Employee.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            try
            {
                connect.Open();
                SqlCommand add = new SqlCommand("Employee_Insert", connect);

                add.CommandType = CommandType.StoredProcedure;
                add.Parameters.AddWithValue("@Surname_Employee", Surname.Text);
                add.Parameters.AddWithValue("@Name_Employee", Name.Text);
                add.Parameters.AddWithValue("@Middle_name_Employee", Middle_name.Text);
                add.Parameters.AddWithValue("@Date_birth", DR.SelectedDate);
                add.Parameters.AddWithValue("@Phone", Phone.Text);
                add.Parameters.AddWithValue("@INN", INN.Text);
                add.Parameters.AddWithValue("@SNILS", SNILS.Text);
                add.Parameters.AddWithValue("@Passport_number_Employee", Passport_number_Employee.Text);
                add.Parameters.AddWithValue("@Passport_series_Employee", Passport_series_Employee.Text);
                add.Parameters.AddWithValue("@Passwd", Passwd.Text);
                add.Parameters.AddWithValue("@Post_ID", Post.SelectedValue);
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
            if (Surname.Text == null || Name.Text == null || Middle_name.Text == null || Phone.Text == null || INN.Text == null ||
                SNILS.Text == null || Passport_number_Employee.Text == null || Passport_series_Employee.Text == null || Passwd.Text == null || Post.Text == null ||
                DR.Text == null)
            { MessageBox.Show("Все поля должны быть заполнены."); return; };
            Regex reg = new Regex(@"[0-9]");
            if (reg.IsMatch(Name.Text) || reg.IsMatch(Surname.Text) || reg.IsMatch(Middle_name.Text))
            { MessageBox.Show("Нельзя вводить числа."); return; }
            Regex reg1 = new Regex(@"[A-Z]");
            Regex reg2 = new Regex(@"[a-z]");
            Regex reg3 = new Regex(@"[А-Я]");
            Regex reg4 = new Regex(@"[а-я]");
            if (reg1.IsMatch(Phone.Text) || reg2.IsMatch(Phone.Text) || reg3.IsMatch(Phone.Text) || reg4.IsMatch(Phone.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            if (reg1.IsMatch(INN.Text) || reg2.IsMatch(INN.Text) || reg3.IsMatch(INN.Text) || reg4.IsMatch(INN.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            if (reg1.IsMatch(SNILS.Text) || reg2.IsMatch(SNILS.Text) || reg3.IsMatch(SNILS.Text) || reg4.IsMatch(SNILS.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            if (reg1.IsMatch(Passport_number_Employee.Text) || reg2.IsMatch(Passport_number_Employee.Text) || reg3.IsMatch(Passport_number_Employee.Text) || reg4.IsMatch(Passport_number_Employee.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            if (reg1.IsMatch(Passport_series_Employee.Text) || reg2.IsMatch(Passport_series_Employee.Text) || reg3.IsMatch(Passport_series_Employee.Text) || reg4.IsMatch(Passport_series_Employee.Text))
            {
                MessageBox.Show("Нельзя вводить буквы."); return;
            }
            DataRowView row = (DataRowView)EmployeeDG.SelectedItem;
            try
            {
                connect.Open();
                SqlCommand upd = new SqlCommand("Employee_Update", connect);

                upd.CommandType = CommandType.StoredProcedure;
                upd.Parameters.AddWithValue("@Surname_Employee", Surname.Text);
                upd.Parameters.AddWithValue("@Name_Employee", Name.Text);
                upd.Parameters.AddWithValue("@Middle_name_Employee", Middle_name.Text);
                upd.Parameters.AddWithValue("@Date_birth", DR.SelectedDate);
                upd.Parameters.AddWithValue("@Phone", Phone.Text);
                upd.Parameters.AddWithValue("@INN", INN.Text);
                upd.Parameters.AddWithValue("@SNILS", SNILS.Text);
                upd.Parameters.AddWithValue("@Passport_number_Employee", Passport_number_Employee.Text);
                upd.Parameters.AddWithValue("@Passport_series_Employee", Passport_series_Employee.Text);
                upd.Parameters.AddWithValue("@Passwd", Passwd.Text);
                upd.Parameters.AddWithValue("@Post_ID", Post.SelectedValue);
                upd.Parameters.AddWithValue("ID_Employee", (int)row["ID_Employee"]);
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
            if (EmployeeDG.SelectedItem == null) { MessageBox.Show("Выберите поле которое хотите удалить"); return; }
            if (EmployeeDG.SelectedIndex == EmployeeDG.Items.Count - 1) { MessageBox.Show("Вы выбрали пустое поле."); return; }
            DataRowView row = (DataRowView)EmployeeDG.SelectedItem;
            try
            {

                connect.Open();
                SqlCommand Del = new SqlCommand("Employee_Delete", connect);
                Del.CommandType = CommandType.StoredProcedure;
                Del.Parameters.AddWithValue("ID_Employee", (int)row["ID_Employee"]);
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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            connect.Open();

            SqlCommand command = new SqlCommand("SELECT ID_Employee, Surname_Employee as 'Фамилия', " +
                "[Name_Employee] AS 'Имя', Middle_name_Employee AS 'Отчество', Name_Post as 'Должность', Phone as 'Номер телефона', INN as 'ИНН', SNILS as 'СНИЛС', Passport_number_Employee as 'Номер паспорта', Passport_series_Employee as 'Серия паспорта', Passwd as 'Пароль', Date_birth as 'Дата рождения' from Employee " +
                "join [dbo].[Post] on [ID_Post]=[Post_ID]", connect);

            DataTable datatbl = new DataTable();

            datatbl.Load(command.ExecuteReader());

            EmployeeDG.ItemsSource = datatbl.DefaultView;

            SqlCommand command1 = new SqlCommand("SELECT ID_Post, Name_Post from [dbo].[Post]", connect);
            DataTable datatbl1 = new DataTable();
            datatbl1.Load(command1.ExecuteReader());
            Post.ItemsSource = datatbl1.DefaultView;
            Post.DisplayMemberPath = "Name_Post";
            Post.SelectedValuePath = "ID_Post";


            connect.Close();
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            connect.Open();

            SqlCommand command = new SqlCommand("SELECT ID_Employee, Surname_Employee as 'Фамилия', " +
                "[Name_Employee] AS 'Имя', Middle_name_Employee AS 'Отчество', Name_Post as 'Должность', Phone as 'Номер телефона', INN as 'ИНН', SNILS as 'СНИЛС', Passport_number_Employee as 'Номер паспорта', Passport_series_Employee as 'Серия паспорта', Passwd as 'Пароль', Date_birth as 'Дата рождения' from Employee " +
                "join [dbo].[Post] on [ID_Post]=[Post_ID]", connect);

            DataTable datatbl = new DataTable();

            datatbl.Load(command.ExecuteReader());

            EmployeeDG.ItemsSource = datatbl.DefaultView;
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

            EmployeeDG.ItemsSource = readFile(openFileDialog.FileName);
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
