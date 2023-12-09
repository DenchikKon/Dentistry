using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Xceed.Wpf.Toolkit;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Runtime.Remoting.Messaging;
using LiveCharts.Wpf.Charts.Base;
using LiveCharts.Wpf;
using LiveCharts;

namespace WPFPractika
{
    public partial class MainWindow : System.Windows.Window
    {
        Regex onlyNumber = new Regex("[^0-9]+");
        Regex onlyLetter = new Regex("[^А-Яа-я]+");

        string query;
        int changeDoctor = -1;
        int changePriceList = -1;
        int changePatient = -1;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            DBManager.LoadData(gridDoctor, DBManager.mainDoctorQuery);
            DBManager.LoadData(gridPriceList, DBManager.mainPriceListQuery);
            DBManager.LoadData(gridPatient, DBManager.mainPatientQuery);
            DBManager.LoadData(gridRecord, DBManager.mainRecordQuery);
            Methods.LoadChart(chart);
        }
        private void OnAutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            PropertyDescriptor propertyDescriptor = (PropertyDescriptor)e.PropertyDescriptor;
            e.Column.Header = propertyDescriptor.DisplayName;
            if (propertyDescriptor.DisplayName == "Id")
            {
                e.Cancel = true;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {            
            Regex Isnumber = new Regex(@"^(80|375)\((44|29|33|25)\)[0-9]{7}$");
           
                if (entryNameDoctor.Text != string.Empty && entrySurnameDoctor.Text != string.Empty && entryPatronymicDoctor.Text != string.Empty
                    && entrySpetializationDoctor.Text != string.Empty && entryPhoneDoctor.Text != string.Empty && entryBirthdayDoctor.SelectedDate.HasValue)
                {
                    if (Isnumber.IsMatch(entryPhoneDoctor.Text))
                    {
                        if (entryBirthdayDoctor.SelectedDate.Value < DateTime.Now)
                        {
                            if (DateTime.Now.Year - entryBirthdayDoctor.SelectedDate.Value.Year >= 18)
                            {
                                if (changeDoctor < 0)
                                {
                                    query = $"Insert Into Doctor values(N'{entryNameDoctor.Text}',N'{entrySurnameDoctor.Text}',N'{entryPatronymicDoctor.Text}'," +
                                    $"'{entryBirthdayDoctor.SelectedDate.Value}','{entryPhoneDoctor.Text}',N'{entrySpetializationDoctor.Text}')";
                                    DBManager.ExecuteQuery(query);
                                    DBManager.LoadData(gridDoctor, DBManager.mainDoctorQuery);                                   
                                }
                                else
                                {
                                query = $"Update Doctor Set Name=N'{entryNameDoctor.Text}',Surname=N'{entrySurnameDoctor.Text}'," +
                                    $" Patronymic=N'{entryPatronymicDoctor.Text}',DateOfBirthDay='{entryBirthdayDoctor.SelectedDate.Value.ToString("yyyy.MM.dd")}'," +
                                    $" Phone='{entryPhoneDoctor.Text}',Spetialization=N'{entrySpetializationDoctor.Text}' where id = {changeDoctor}";
                                DBManager.ExecuteQuery(query);
                                DBManager.LoadData(gridDoctor, DBManager.mainDoctorQuery);
                                changeDoctor = -1;
                                }
                            entryNameDoctor.Text = string.Empty;
                            entrySurnameDoctor.Text = string.Empty;
                            entryPhoneDoctor.Text = string.Empty;
                            entryPatronymicDoctor.Text = string.Empty;
                            entrySpetializationDoctor.Text = string.Empty;
                        }
                            else
                            Xceed.Wpf.Toolkit.MessageBox.Show("Для устройства на работу требуется 18 лет");
                        }
                        else
                        Xceed.Wpf.Toolkit.MessageBox.Show("Дата рождения не может превышать текущую дату");
                    }
                    else
                    Xceed.Wpf.Toolkit.MessageBox.Show("Введите кореектные данные в поле телефон (80(00)0000000)");
                }
                else
                Xceed.Wpf.Toolkit.MessageBox.Show("Требуется заполнить все поля");
           
          
        }

        private void SavePriceList_Click(object sender, RoutedEventArgs e)
        {
            
                if (entryTitlePriceList.Text != string.Empty && decimal.TryParse(entryPricePriceList.Text, out decimal value))
                {
                    if (changePriceList < 0)
                    {
                    query = $"Insert into PriceList values(N'{entryTitlePriceList.Text}','{value}')";
                    DBManager.ExecuteQuery(query);
                    DBManager.LoadData(gridPriceList, DBManager.mainPriceListQuery);
                    }
                    else
                    {
                    query = $"Update PriceList Set Title=N'{entryTitlePriceList.Text}',Price={value} where id={changePriceList}";
                    DBManager.ExecuteQuery(query);
                    DBManager.LoadData(gridPriceList, DBManager.mainPriceListQuery);
                    changePriceList = -1;                   
                    }
                    entryTitlePriceList.Text = string.Empty;
                    entryPricePriceList.Text = string.Empty;
            }
                else
                Xceed.Wpf.Toolkit.MessageBox.Show("Требуется заполнить все поля. Поле цены должно быть числовым с точкой");
           
        }

        private void Window_Initialized(object sender, EventArgs e)
        {
   
          
        }
        private void savePatient_Click(object sender, RoutedEventArgs e)
        {
            Regex Isnumber = new Regex(@"^(80|375)\((44|29|33|25)\)[0-9]{7}$");
           
                if (entryNamePatient.Text != string.Empty && EntrySurnamePatient.Text != string.Empty && entryPatronymicPatient.Text != string.Empty
                    && entryAddressPatient.Text != string.Empty && entryPhonePatient.Text != string.Empty && entryBirthDayPatient.SelectedDate.HasValue)
                {
                    if (Isnumber.IsMatch(entryPhonePatient.Text))
                    {
                        if (entryBirthDayPatient.SelectedDate.Value < DateTime.Now)
                        {
                            if (changePatient < 0)
                            {
                                query = $"Insert Into Patient values(N'{entryNamePatient.Text}',N'{EntrySurnamePatient.Text}',N'{entryPatronymicPatient.Text}',N'{entryAddressPatient.Text}'," +
                                $"'{entryBirthDayPatient.SelectedDate.Value}','{entryPhonePatient.Text}')";
                                DBManager.ExecuteQuery(query);
                                DBManager.LoadData(gridPatient, DBManager.mainPatientQuery);
                                
                            }
                            else
                            {
                            query = $"Update Patient Set Name=N'{entryNamePatient.Text}',Surname=N'{EntrySurnamePatient.Text}',Patronymic=N'{entryPatronymicPatient.Text}',Address=N'{entryAddressPatient.Text}'," +
                            $"DateOfBirthDay='{entryBirthDayPatient.SelectedDate.Value}',Phone='{entryPhonePatient.Text}' where Id={changePatient}";
                            DBManager.ExecuteQuery(query);
                            DBManager.LoadData(gridPatient, DBManager.mainPatientQuery);
                            Methods.LoadChart(chart);
                            changePatient = -1;
                            }
                            entryNamePatient.Text = string.Empty;
                            EntrySurnamePatient.Text = string.Empty;
                            entryPhonePatient.Text = string.Empty;
                            entryPatronymicPatient.Text = string.Empty;
                            entryAddressPatient.Text = string.Empty;
                    }
                        else
                            Xceed.Wpf.Toolkit.MessageBox.Show("Дата рождения не может превышать текущую дату");
                    }
                    else
                         Xceed.Wpf.Toolkit.MessageBox.Show("Введите кореектные данные в поле телефон (80(00)0000000)");
                }
                else
                    Xceed.Wpf.Toolkit.MessageBox.Show("Требуется заполнить все поля");
           
        }

        private void DeletePatient_Click(object sender, RoutedEventArgs e)
        {
            try 
            { 
            DataRowView row = (DataRowView)gridPatient.SelectedItem;
            query = $"Delete From Patient where id={row[0]}";
            DBManager.ExecuteQuery(query);
            DBManager.LoadData(gridPatient, DBManager.mainPatientQuery);
            }
            catch (Exception) 
            {
                Xceed.Wpf.Toolkit.MessageBox.Show("Невозможно удалить данную строку");
            }
        }

        private void DeletePriceList_Click(object sender, RoutedEventArgs e)
        {
            try 
            { 
            DataRowView row = (DataRowView)gridPriceList?.SelectedItem;
            query = $"Delete From PriceList where id={row[0]}";
            DBManager.ExecuteQuery(query);
            DBManager.LoadData(gridPriceList, DBManager.mainPriceListQuery);
            }
            catch (Exception)
            {
                Xceed.Wpf.Toolkit.MessageBox.Show("Невозможно удалить данную строку");
            }
            finally
            {
                if(DBManager.DentistryDBConnetion.State == ConnectionState.Open)
                    DBManager.ConnectClose();
            }
        }

        private void deleteDoctor_Click(object sender, RoutedEventArgs e)
        {
            try 
            { 
            DataRowView row = (DataRowView)gridDoctor.SelectedItem;
            query = $"Delete From Doctor where id={row[0]}";
            DBManager.ExecuteQuery(query);
            DBManager.LoadData(gridDoctor, DBManager.mainDoctorQuery);
            }
            catch (Exception)
            {
                Xceed.Wpf.Toolkit.MessageBox.Show("Невозможно удалить данную строку");
            }
            finally
            {
                if (DBManager.DentistryDBConnetion.State == ConnectionState.Open)
                    DBManager.ConnectClose();
            }
        }

        private void UpdateDoctor_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)gridDoctor.SelectedItem;
                changeDoctor = int.Parse(row[0].ToString());
                string[] splited = row[1].ToString().Split(' ');
                entryNameDoctor.Text = splited[0];
                entrySurnameDoctor.Text = splited[1];
                entryPatronymicDoctor.Text = splited[2];
                entrySpetializationDoctor.Text = row[2].ToString();
                entryPhoneDoctor.Text = row[3].ToString().Trim();
                entryBirthdayDoctor.Text = row[4].ToString();
            }
            catch (Exception){ Xceed.Wpf.Toolkit.MessageBox.Show("Невозможно изменить данную строку"); }
            finally
            {
                if (DBManager.DentistryDBConnetion.State == ConnectionState.Open)
                    DBManager.ConnectClose();
            }
        }

        private void UpdatePriceList_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)gridPriceList.SelectedItem;
                changePriceList = int.Parse(row[0].ToString());
                entryTitlePriceList.Text = row[1].ToString().Trim();
                entryPricePriceList.Text = row[2].ToString().Trim();
            }
            catch (Exception)
            {
                Xceed.Wpf.Toolkit.MessageBox.Show("Невозможно изменить данную строку");
            }
            finally
            {
                if (DBManager.DentistryDBConnetion.State == ConnectionState.Open)
                    DBManager.ConnectClose();
            }
        }
        private void UpdatePatient_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)gridPatient.SelectedItem;
                changePatient = int.Parse(row[0].ToString());
                string[] splited = row[1].ToString().Split(' ');
                entryNamePatient.Text = splited[0];
                EntrySurnamePatient.Text = splited[1];
                entryPatronymicPatient.Text = splited[2];
                entryAddressPatient.Text = row[2].ToString().Trim();
                entryPhonePatient.Text = row[3].ToString().Trim();
                entryBirthDayPatient.Text = row[4].ToString().Trim();
            }
            catch (Exception) { Xceed.Wpf.Toolkit.MessageBox.Show("Невозможно изменить данную строку"); }
            finally
            {
                if (DBManager.DentistryDBConnetion.State == ConnectionState.Open)
                    DBManager.ConnectClose();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            AddRecord addRecord = new AddRecord();
            addRecord.Owner = this;
            addRecord.ShowDialog();
            DBManager.LoadData(gridRecord, DBManager.mainRecordQuery);
            Methods.LoadChart(chart);
        }
        private void exportExcelDoctor_Click(object sender, RoutedEventArgs e)
        {
            Methods.ExportExcel(gridDoctor,"Врачи");
        }


        private void exportExcelPriceList_Click(object sender, RoutedEventArgs e)
        {
            Methods.ExportExcel(gridPriceList,"Услуги");
        }

        private void exportExcelPatient_Click(object sender, RoutedEventArgs e)
        {
            Methods.ExportExcel(gridPatient,"Клиенты");
        }

        private void exportExcelRecord_Click(object sender, RoutedEventArgs e)
        {
            Methods.ExportExcel(gridRecord, "Записи");
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Methods.Search(gridRecord, searchRecord);
        }

        private void SearchPatient_TextChanged(object sender, TextChangedEventArgs e)
        {
            Methods.Search(gridPatient, searchPatient);
        }

        private void SearchPriceList_TextChanged(object sender, TextChangedEventArgs e)
        {
            Methods.Search(gridPriceList, searchPriceList);
        }

        private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {
            Methods.Search(gridDoctor, searchDoctor);
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            
        }
        private void Row_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            DataRowView row = (DataRowView)gridRecord.SelectedItems[0];
            RecordInfo recordInfo = new RecordInfo();
            recordInfo.Owner = this;
            recordInfo.CheckDoctor.Content = row[2].ToString();
            recordInfo.CheckPatient.Content = row[1].ToString();
            recordInfo.checkDate.Content = row[3].ToString();
            recordInfo.checkTime.Content = row[4].ToString();
            recordInfo.checkFullPrice.Content = row[5].ToString();
            query = $"Select PriceList.Title, PriceList.Price ,ListOfServices.Count From ListOfServices " +
                $"inner join PriceList on PriceList.Id = ListOfServices.IdPriceList inner join Record on Record.Id = ListOfServices.idRecord where Record.Id = {row[0].ToString()}";
            DBManager.LoadData(recordInfo.gridPriceListInfo,query);
            recordInfo.Show();
        }
        private void exportWord_Click(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)gridRecord.SelectedItems[0];
            var wordApp = new Word.Application();
            var doc = wordApp.Documents.Open(@"D:\КППрактика\WPFPractika\чек1.docx");
            // Заменяем поля в документе на нужные значения
            doc.Content.Find.Execute(FindText: "{date}", ReplaceWith: row["Дата"].ToString());
            doc.Content.Find.Execute(FindText: "{workerName}", ReplaceWith: row["Доктор"].ToString());
            doc.Content.Find.Execute(FindText: "{client}", ReplaceWith: row["Пациент"].ToString());
            doc.Content.Find.Execute(FindText: "{fullPrice}", ReplaceWith: row["Стоимость"].ToString());
            doc.Content.Find.Execute(FindText: "{time}", ReplaceWith: row["Время"].ToString());
            System.Data.DataTable dataTable = new System.Data.DataTable();
            // Получаем таблицу из документа
            DBManager.ConnectOpen();
            query = $"Select PriceList.Title, PriceList.Price ,ListOfServices.Count from ListOfServices inner join PriceList on PriceList.Id = ListOfServices.IdPriceList" +
                $" where ListOfServices.idRecord = {row[0]}";
            SqlCommand command = new SqlCommand(query, DBManager.DentistryDBConnetion);
            SqlDataReader reader = command.ExecuteReader();
            if (reader.HasRows)
            {
                object objMissing = System.Reflection.Missing.Value;
                int rowIndex = 2;
                while (reader.Read())
                {
                    // Добавляем строки в таблицу
                    doc.Tables[1].Rows.Add(ref objMissing);
                    doc.Tables[1].Cell(rowIndex, 1).Range.Text = reader.GetString(0);
                    doc.Tables[1].Cell(rowIndex, 2).Range.Text = reader.GetDecimal(1).ToString();
                    doc.Tables[1].Cell(rowIndex, 3).Range.Text = reader.GetInt32(2).ToString();
                    rowIndex++;
                }
            }
            reader.Close();
            DBManager.ConnectClose();
                wordApp.Visible = true;
        }
        private void DeleteRecord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)gridRecord.SelectedItem;
                query = $"Delete From ListOfServices where idRecord={row[0]}";
                DBManager.ExecuteQuery(query);
                query = $"Delete From Record where id={row[0]}";
                DBManager.ExecuteQuery(query);
                DBManager.LoadData(gridRecord, DBManager.mainRecordQuery);
            }
            catch (Exception)
            {
                Xceed.Wpf.Toolkit.MessageBox.Show("Невозможно удалить данную строку");
            }
            finally
            {
                if (DBManager.DentistryDBConnetion.State == ConnectionState.Open)
                    DBManager.ConnectClose();
            }
        }

        private void entryNamePatient_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = onlyLetter.IsMatch(e.Text);
        }

        private void EntrySurnamePatient_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = onlyLetter.IsMatch(e.Text);
        }

        private void entryPatronymicPatient_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = onlyLetter.IsMatch(e.Text);
        }

        private void entryNameDoctor_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = onlyLetter.IsMatch(e.Text);
        }

        private void entrySurnameDoctor_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = onlyLetter.IsMatch(e.Text);
        }

        private void entryPatronymicDoctor_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = onlyLetter.IsMatch(e.Text);
        }

        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = onlyLetter.IsMatch(e.Text);
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (filterName.Text != string.Empty || filterDateEnd.SelectedDate.HasValue || filterDateStart.SelectedDate.HasValue
                || filterPriceStart.Text != string.Empty || filterPriceEnd.Text != string.Empty)
            {
                string startDate = filterDateStart.SelectedDate.HasValue ? filterDateStart.SelectedDate.Value.ToString("yyyy.MM.dd") : "0001.01.01";
                string endDate = filterDateEnd.SelectedDate.HasValue ? filterDateStart.SelectedDate.Value.ToString("yyyy.MM.dd") : "9999.12.12";
                decimal startValue = decimal.TryParse(filterPriceStart.Text, out decimal firstvalue) ? firstvalue : 0;
                decimal lastValue = decimal.TryParse(filterPriceEnd.Text, out decimal secondvalue) ? secondvalue : decimal.MaxValue;
                if (startValue<lastValue)
                {
                    query = DBManager.mainRecordQuery + $" where Concat(Trim(Patient.Name),' ',Trim(Patient.Surname),' ',Trim(Patient.Patronymic)) Like N'%{filterName.Text}%' AND " +
                        $"Record.Date between '{startDate}' and '{endDate}' AND Record.FullPrice between {startValue} and {lastValue}";
                    DBManager.LoadData(gridRecord, query);
                }
                else
                    Xceed.Wpf.Toolkit.MessageBox.Show("Начальная значение не может быть больше конечной");
            }
            else
               Xceed.Wpf.Toolkit.MessageBox.Show("Требуется что бы хотя бы одно поле было заполнено");
        }

        private void TextBox_PreviewTextInput_1(object sender, TextCompositionEventArgs e)
        {
            e.Handled = onlyNumber.IsMatch(e.Text);
        }

        private void TextBox_PreviewTextInput_2(object sender, TextCompositionEventArgs e)
        {
            e.Handled = onlyNumber.IsMatch(e.Text);
            
        }

        private void resetFilters_Click(object sender, RoutedEventArgs e)
        {
            DBManager.LoadData(gridRecord, DBManager.mainRecordQuery);
            filterName.Text = string.Empty;
            filterPriceEnd.Text = string.Empty;
            filterPriceStart.Text = string.Empty;
            filterDateStart.Text = null;
            filterDateEnd.Text = null;
        }
    }
}
